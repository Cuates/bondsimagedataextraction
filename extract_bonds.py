import os
import glob
import json
import time
import random
import mimetypes
from concurrent.futures import ThreadPoolExecutor, as_completed
import pandas as pd
from tqdm import tqdm
from pydantic import BaseModel, Field
from google import genai
from google.genai import types, errors

# 1. Define the Pydantic schema for structured output
class BondDetails(BaseModel):
    serial_number: str = Field(description="Unique bond ID, usually bottom right or corners")
    series: str = Field(description="Series of the bond, e.g., EE, I, E")
    denomination: str = Field(description="Face value of the bond, e.g., $50, $100")
    issue_date: str = Field(description="Issue date, usually formatted MM/YYYY")
    owner_names: str = Field(description="The full registered names printed on the bond")
    co_owner: str = Field(description="The co-owner printed on the bond, if present. Return empty string if not found.", default="")
    address: str = Field(description="The address printed on the bond, if present. Return empty string if not found.", default="")
    notes: str = Field(description="Any other data worth noting from the bond to track")

def process_image(image_path: str, client: genai.Client) -> dict:
    """Processes a single image and extracts bond details."""
    result = {
        "file_name": os.path.basename(image_path),
        "serial_number": None,
        "series": None,
        "denomination": None,
        "issue_date": None,
        "owner_names": None,
        "co_owner": None,
        "address": None,
        "notes": None,
        "status": "Success"
    }

    try:
        # Determine mime type
        mime_type, _ = mimetypes.guess_type(image_path)
        if not mime_type:
            if image_path.lower().endswith('.heic'):
                mime_type = "image/heic"
            else:
                mime_type = "image/jpeg"

        # Read image bytes
        with open(image_path, "rb") as f:
            image_bytes = f.read()

        image_part = types.Part.from_bytes(
            data=image_bytes,
            mime_type=mime_type
        )

        prompt = (
            "Analyze this image of a US Savings Bond. Please extract the requested details. "
            "Note that the image might be a photo taken from an iPhone and could be skewed, upside down, or sideways. "
            "Adjust your reading direction accordingly to accurately extract the text."
        )

        max_retries = 5
        for attempt in range(max_retries):
            try:
                response = client.models.generate_content(
                    model='gemini-3.1-flash-lite', # Try 'gemini-2.5-flash' or 'gemini-1.5-flash' if 3.5 continues to fail
                    contents=[image_part, prompt],
                    config=types.GenerateContentConfig(
                        response_mime_type="application/json",
                        response_schema=BondDetails,
                        temperature=0.1,
                    ),
                )
                break # Success! Exit the retry loop.
            except errors.APIError as e:
                # If we get a 503 (Unavailable) or 429 (Rate Limit), wait and try again
                if ("503" in str(e) or "429" in str(e)) and attempt < max_retries - 1:
                    time.sleep((2 ** attempt) + random.uniform(0.5, 1.5)) # Exponential backoff: 1s, 2s, 4s, 8s + jitter
                    continue
                raise # If it's a different error or we are out of retries, fail properly

        # Safely extract structured response
        if hasattr(response, 'parsed') and response.parsed is not None:
            extracted = response.parsed.model_dump()
        else:
            extracted = json.loads(response.text)

        result.update(extracted)

    except (errors.APIError, FileNotFoundError, json.JSONDecodeError, ValueError, OSError) as e:
        result["status"] = f"Error: {type(e).__name__} - {str(e)}"

    return result

def main():
    print("========================================")
    print(" US Savings Bond Extractor")
    print("========================================")

    tier_choice = input("Are you using the Free Tier API? (y/n): ").strip().lower()
    is_free_tier = tier_choice == 'y' or tier_choice == 'yes'

    # Setup paths
    input_folder = input("Enter the path to the folder containing bond images: ").strip()
    # Strip quotes if dragged-and-dropped in terminal
    input_folder = input_folder.strip('"').strip("'")

    if not os.path.isdir(input_folder):
        print(f"Error: Directory '{input_folder}' does not exist.")
        return

    output_file = input("Enter the desired output Excel file name (e.g., bonds_data.xlsx): ").strip()
    if not output_file.lower().endswith('.xlsx'):
        output_file += '.xlsx'

    # Initialize Gemini client
    try:
        client = genai.Client()
    except (errors.APIError, ValueError, KeyError, EnvironmentError) as e:
        print(f"Failed to initialize Gemini Client. Make sure the GEMINI_API_KEY environment variable is set.\nError: {e}")
        return

    # Find images (supporting common formats)
    supported_extensions = ('*.png', '*.jpg', '*.jpeg', '*.heic')
    image_paths = []
    for ext in supported_extensions:
        # Check both lowercase and uppercase extensions
        image_paths.extend(glob.glob(os.path.join(input_folder, ext)))
        image_paths.extend(glob.glob(os.path.join(input_folder, ext.upper())))

    # Remove duplicates and sort alphabetically so they process in a predictable order
    image_paths = sorted(list(set(image_paths)))

    if not image_paths:
        print(f"No images found in '{input_folder}'. Supported formats: png, jpg, jpeg, heic.")
        return

    # Check for existing Excel file to resume progress
    existing_results = []
    if os.path.exists(output_file):
        try:
            existing_df = pd.read_excel(output_file)
            # Replace NaNs with None for cleaner dictionaries
            existing_df = existing_df.where(pd.notna(existing_df), None)

            for _, row in existing_df.iterrows():
                row_dict = row.to_dict()
                status = str(row_dict.get('status', ''))
                # We consider an image processed if it didn't fail due to Quota Exhaustion
                if "429" not in status and "RESOURCE_EXHAUSTED" not in status:
                    existing_results.append(row_dict)

            already_processed_names = {r['file_name'] for r in existing_results}

            original_count = len(image_paths)
            image_paths = [p for p in image_paths if os.path.basename(p) not in already_processed_names]

            if original_count > len(image_paths):
                print(f"\n[i] Found existing '{os.path.basename(output_file)}'.")
                print(f"[i] Resuming progress: skipping {original_count - len(image_paths)} already processed images.")

        except Exception as e:
            print(f"\n[!] Warning: Could not read existing Excel file to resume progress. Starting fresh.\nError: {e}")
            existing_results = []

    if not image_paths:
        print("\nAll images in the folder have already been processed! No further extraction needed.")
        return

    print(f"\nFound {len(image_paths)} new images to process. Starting extraction...")
    print(f"Mode: {'Sequential (Free Tier)' if is_free_tier else 'Concurrent (Paid Tier)'}\n")

    # Start with existing results so they are included in the final Excel save
    results = existing_results.copy()

    if is_free_tier:
        # Process one by one with a progress bar and forced sleep
        for path in tqdm(image_paths, desc="Processing Bonds (Free Tier)", unit="img"):
            result = process_image(path, client)

            # Check for API Quota exhaustion to stop early
            if "429" in result["status"] and "RESOURCE_EXHAUSTED" in result["status"]:
                print(f"\n[!] Quota Exhausted (429 Error) encountered on {os.path.basename(path)}.")
                print("[!] Stopping process early to prevent unnecessary delays and preserve already extracted data.")
                break

            results.append(result)
            # Sleep for 4.1 seconds to stay strictly under the 15 requests per minute Free Tier limit
            time.sleep(4.1)
    else:
        # Process images concurrently
        max_workers = 15
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            # Submit all tasks
            future_to_path = {executor.submit(process_image, path, client): path for path in image_paths}

            # Process results as they complete, with a progress bar
            for future in tqdm(as_completed(future_to_path), total=len(image_paths), desc="Processing Bonds (Paid Tier)", unit="img"):
                result = future.result()

                # Check for API Quota exhaustion to stop early
                if "429" in result["status"] and "RESOURCE_EXHAUSTED" in result["status"]:
                    print(f"\n[!] Quota Exhausted (429 Error) encountered on {result['file_name']}.")
                    print("[!] Stopping process early to prevent unnecessary delays and preserve already extracted data.")

                    # Attempt to cancel any futures that haven't started running yet
                    for f in future_to_path:
                        f.cancel()
                    break

                results.append(result)

    # Do not output an empty Excel file
    if not results:
        print("\nNo data was successfully extracted. The Excel file will NOT be generated.")
        return

    # Save to Excel
    print(f"\nSaving {len(results)} extracted records to Excel...")
    df = pd.DataFrame(results)

    # Reorder columns to ensure file_name and status are at intuitive positions
    cols = ['file_name', 'serial_number', 'series', 'denomination', 'issue_date', 'owner_names', 'co_owner', 'address', 'notes', 'status']
    df = df.reindex(columns=[c for c in cols if c in df.columns])

    # Sort rows alphabetically by file name before saving
    df.sort_values(by='file_name', inplace=True)

    try:
        df.to_excel(output_file, index=False)
        print(f"Done! Extracted data saved to: {os.path.abspath(output_file)}")
    except (PermissionError, OSError, ValueError, ImportError) as e:
        print(f"Error saving to Excel: {e}")

if __name__ == "__main__":
    main()
