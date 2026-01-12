from msal import PublicClientApplication
import requests
import csv
import os
from datetime import datetime
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# === CONFIG ===
# Using Microsoft Graph PowerToys client ID (pre-registered public client)
CLIENT_ID = os.getenv("CLIENT_ID")  # Microsoft Graph PowerToys
SCOPES = ["Notes.Read"]
AUTHORITY = "https://login.microsoftonline.com/common"  # Works for any tenant

# If you want to limit to a single notebook, set this to its display name.
# Example: TARGET_NOTEBOOK = "Notitieblok van Joris"
# If you set it to None, it will list sections & page counts for ALL notebooks.
TARGET_NOTEBOOK = "Notitieblok van Joris"

# Target section groups to process
TARGET_SECTION_GROUPS = ["2022 SNs", "2023 SNs", "2024 SNs"]

# Where to store the CSV summary
EXPORT_DIR = "onenote_export"
CSV_FILENAME = f"section_page_counts_{datetime.now().strftime('%Y%m%d')}.csv"


def acquire_token():
    """Authenticate and return an access token."""
    app = PublicClientApplication(
        CLIENT_ID,
        authority=AUTHORITY
    )

    print("🔐 Starting authentication...")
    result = app.acquire_token_interactive(scopes=SCOPES)

    if "access_token" in result:
        print("✅ Authentication successful!")
        return result["access_token"]

    print("❌ Authentication failed!")
    print("Error details:", result.get("error"))
    print("Error description:", result.get("error_description"))
    raise SystemExit(1)


def get_all_pages_for_section(section, headers):
    """
    Given a section object from the Graph response, fetch all its pages
    (handling pagination) and return the list.
    """
    pages_url = section.get("pagesUrl")
    if not pages_url:
        return []

    pages_response = requests.get(pages_url, headers=headers)
    pages_response.raise_for_status()
    pages_data = pages_response.json()

    all_pages = pages_data.get("value", [])
    next_link = pages_data.get("@odata.nextLink")

    while next_link:
        next_response = requests.get(next_link, headers=headers)
        next_response.raise_for_status()
        next_data = next_response.json()

        batch_pages = next_data.get("value", [])
        all_pages.extend(batch_pages)

        next_link = next_data.get("@odata.nextLink")

    return all_pages


def main():
    access_token = acquire_token()
    headers = {"Authorization": f"Bearer {access_token}"}

    # Prepare export dir and CSV path
    os.makedirs(EXPORT_DIR, exist_ok=True)
    csv_path = os.path.join(EXPORT_DIR, CSV_FILENAME)

    # Collect rows for CSV: one row per section
    csv_rows = []

    print("📡 Fetching notebooks...")
    notebooks_response = requests.get(
        "https://graph.microsoft.com/v1.0/me/onenote/notebooks",
        headers=headers,
    )
    notebooks_response.raise_for_status()
    notebooks = notebooks_response.json().get("value", [])

    if not notebooks:
        print("⚠️ No notebooks found for this account.")
        return

    print(f"📒 Found {len(notebooks)} notebook(s).")

    # Filter notebooks if TARGET_NOTEBOOK is set
    if TARGET_NOTEBOOK:
        notebooks = [nb for nb in notebooks if nb.get("displayName") == TARGET_NOTEBOOK]
        if not notebooks:
            print(f"❌ Notebook '{TARGET_NOTEBOOK}' not found.")
            print("Available notebooks:")
            for nb in notebooks:
                print(f"  - {nb.get('displayName')}")
            return

    for notebook in notebooks:
        notebook_name = notebook.get("displayName")
        print(f"\n==============================")
        print(f"📓 Notebook: {notebook_name}")

        # Get section groups from the notebook
        section_groups_url = notebook.get("sectionGroupsUrl")
        if not section_groups_url:
            print("  ⚠️ No sectionGroupsUrl found for this notebook.")
            continue

        section_groups_response = requests.get(section_groups_url, headers=headers)
        section_groups_response.raise_for_status()
        all_section_groups = section_groups_response.json().get("value", [])

        # Handle pagination for section groups
        next_link = section_groups_response.json().get("@odata.nextLink")
        while next_link:
            next_response = requests.get(next_link, headers=headers)
            next_response.raise_for_status()
            next_data = next_response.json()
            all_section_groups.extend(next_data.get("value", []))
            next_link = next_data.get("@odata.nextLink")

        if not all_section_groups:
            print("  ⚠️ No section groups found in this notebook.")
            continue

        print(f"  📁 Found {len(all_section_groups)} section group(s) in this notebook.")

        # Filter for target section groups
        target_section_groups = [
            sg for sg in all_section_groups
            if sg.get("displayName") in TARGET_SECTION_GROUPS
        ]

        if not target_section_groups:
            print(f"  ⚠️ None of the target section groups {TARGET_SECTION_GROUPS} found.")
            print("  Available section groups:")
            for sg in all_section_groups:
                print(f"    - {sg.get('displayName')}")
            continue

        print(f"  ✅ Found {len(target_section_groups)} target section group(s)")

        # Process each target section group
        for section_group in target_section_groups:
            section_group_name = section_group.get("displayName")
            print(f"\n  📁 Section Group: {section_group_name}")

            # Get sections within this section group
            sections_url = section_group.get("sectionsUrl")
            if not sections_url:
                print("    ⚠️ No sectionsUrl found for this section group.")
                continue

            sections_response = requests.get(sections_url, headers=headers)
            sections_response.raise_for_status()
            sections_data = sections_response.json()
            sections = sections_data.get("value", [])

            # Handle pagination for sections
            next_link = sections_data.get("@odata.nextLink")
            while next_link:
                next_response = requests.get(next_link, headers=headers)
                next_response.raise_for_status()
                next_data = next_response.json()
                sections.extend(next_data.get("value", []))
                next_link = next_data.get("@odata.nextLink")

            if not sections:
                print("    ⚠️ No sections found in this section group.")
                continue

            print(f"    📂 Found {len(sections)} section(s) in this section group.")

            # Process each section
            for section in sections:
                section_name = section.get("displayName") or ""
                print(f"\n    ➤ Section: {section_name}")

                try:
                    all_pages = get_all_pages_for_section(section, headers)
                    page_count = len(all_pages)
                    print(f"      📄 Page count: {page_count}")

                    # Add a CSV row for this section
                    csv_rows.append(
                        {
                            "notebook_name": notebook_name,
                            "notebook_id": notebook.get("id"),
                            "section_group_name": section_group_name,
                            "section_group_id": section_group.get("id"),
                            "section_name": section_name,
                            "section_id": section.get("id"),
                            "page_count": page_count,
                        }
                    )
                except requests.exceptions.RequestException as e:
                    print(f"      ❌ Failed to fetch pages for section '{section_name}': {e}")
                    if hasattr(e, "response") and e.response is not None:
                        print(f"        Status: {e.response.status_code}")
                        print(f"        Response: {e.response.text}")

    # Write CSV if we have any data
    if csv_rows:
        # Sort by page_count in descending order (highest first)
        csv_rows.sort(key=lambda x: x["page_count"], reverse=True)
        
        fieldnames = [
            "notebook_name",
            "notebook_id",
            "section_group_name",
            "section_group_id",
            "section_name",
            "section_id",
            "page_count",
        ]
        with open(csv_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(csv_rows)

        print(f"\n✅ CSV export written to: {csv_path}")
    else:
        print("\n⚠️ No section data collected; CSV not written.")


if __name__ == "__main__":
    main()


