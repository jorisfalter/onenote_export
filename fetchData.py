from msal import PublicClientApplication
import requests
from bs4 import BeautifulSoup
import os
from urllib.parse import urlparse, unquote
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# === CONFIG ===
# Using Microsoft Graph PowerToys client ID (pre-registered public client)
CLIENT_ID = os.getenv("CLIENT_ID")  # Microsoft Graph PowerToys
SCOPES = ["Notes.Read"]
AUTHORITY = "https://login.microsoftonline.com/common"  # Works for any tenant

# Create export directory
EXPORT_DIR = "onenote_export"
os.makedirs(EXPORT_DIR, exist_ok=True)
os.makedirs(os.path.join(EXPORT_DIR, "images"), exist_ok=True)

# === AUTHENTICATE ===
app = PublicClientApplication(
    CLIENT_ID,
    authority=AUTHORITY
)

print("🔐 Starting authentication...")
result = app.acquire_token_interactive(scopes=SCOPES)

# Check if authentication was successful
if "access_token" in result:
    print("✅ Authentication successful!")
    access_token = result['access_token']
else:
    print("❌ Authentication failed!")
    print("Error details:", result.get('error'))
    print("Error description:", result.get('error_description'))
    exit(1)

headers = {"Authorization": f"Bearer {access_token}"}

def download_image(img_url, img_index, page_title):
    """Download an image from OneNote and save it locally"""
    try:
        print(f"      📥 Downloading image {img_index}...")
        img_response = requests.get(img_url, headers=headers)
        img_response.raise_for_status()
        
        # Create a safe filename
        safe_title = "".join(c for c in page_title if c.isalnum() or c in (' ', '-', '_')).strip()
        safe_title = safe_title.replace(' ', '_')
        
        # Determine file extension from content type or default to jpg
        content_type = img_response.headers.get('content-type', '')
        if 'png' in content_type:
            ext = '.png'
        elif 'gif' in content_type:
            ext = '.gif'
        else:
            ext = '.jpg'  # Default
        
        filename = f"{safe_title}_image_{img_index}{ext}"
        filepath = os.path.join(EXPORT_DIR, "images", filename)
        
        with open(filepath, 'wb') as f:
            f.write(img_response.content)
        
        print(f"        ✅ Saved as: {filename}")
        return filename
        
    except Exception as e:
        print(f"        ❌ Failed to download image: {e}")
        return None

# === FETCH SPECIFIC NOTEBOOK AND SECTION ===
print("📡 Fetching notebooks...")

# Target configuration
TARGET_NOTEBOOK = "Notitieblok van Joris"
TARGET_SECTION = "MADLI"

try:
    notebooks_response = requests.get("https://graph.microsoft.com/v1.0/me/onenote/notebooks", headers=headers)
    notebooks_response.raise_for_status()
    notebooks = notebooks_response.json()

    # Find the target notebook
    target_notebook = None
    for notebook in notebooks.get("value", []):
        if notebook['displayName'] == TARGET_NOTEBOOK:
            target_notebook = notebook
            break
    
    if not target_notebook:
        print(f"❌ Notebook '{TARGET_NOTEBOOK}' not found!")
        print("Available notebooks:")
        for nb in notebooks.get("value", []):
            print(f"  - {nb['displayName']}")
        exit(1)

    print(f"\n📒 Found target notebook: {target_notebook['displayName']}")

    # Get sections from the target notebook
    sections_response = requests.get(target_notebook['sectionsUrl'], headers=headers)
    sections_response.raise_for_status()
    sections = sections_response.json()
    
    # Find the target section
    target_section = None
    for section in sections.get("value", []):
        if section['displayName'] == TARGET_SECTION:
            target_section = section
            break
    
    if not target_section:
        print(f"❌ Section '{TARGET_SECTION}' not found!")
        print("Available sections:")
        for sec in sections.get("value", []):
            print(f"  - {sec['displayName']}")
        exit(1)

    print(f"  📂 Found target section: {target_section['displayName']}")

    # Get all pages from the target section (with pagination)
    pages_response = requests.get(target_section['pagesUrl'], headers=headers)
    pages_response.raise_for_status()
    pages_data = pages_response.json()
    
    all_pages = pages_data.get("value", [])
    print(f"    📄 Found {len(all_pages)} page(s) in first batch")
    
    # Check for pagination (@odata.nextLink)
    next_link = pages_data.get("@odata.nextLink")
    page_batch = 1
    
    while next_link:
        page_batch += 1
        print(f"    📄 Fetching batch {page_batch}...")
        
        next_response = requests.get(next_link, headers=headers)
        next_response.raise_for_status()
        next_data = next_response.json()
        
        batch_pages = next_data.get("value", [])
        all_pages.extend(batch_pages)
        print(f"    📄 Added {len(batch_pages)} more page(s) (total: {len(all_pages)})")
        
        # Update next_link for next iteration
        next_link = next_data.get("@odata.nextLink")
    
    print(f"    📄 Final count: {len(all_pages)} page(s) total in section")
    
    if not all_pages:
        print("    ⚠️ No pages found in this section.")
        exit(1)

    # Process all pages in the section
    for page_idx, page in enumerate(all_pages, 1):
        print(f"\n--- Processing page {page_idx}/{len(all_pages)}: {page['title']} ---")

        content_response = requests.get(page['contentUrl'], headers=headers)
        content_response.raise_for_status()
        html = content_response.text
        
        # Create a safe page title for filenames
        safe_title = "".join(c for c in page['title'] if c.isalnum() or c in (' ', '-', '_')).strip()
        safe_title = safe_title.replace(' ', '_')
        if not safe_title:  # Fallback for empty titles
            safe_title = f"page_{page_idx}"
        
        # Save raw HTML for inspection
        # html_filename = f"{safe_title}_raw.html"
        # with open(os.path.join(EXPORT_DIR, html_filename), "w", encoding="utf-8") as f:
        #     f.write(html)
        # print(f"    💾 Raw HTML saved: {html_filename}")
        
        # Parse and extract text
        soup = BeautifulSoup(html, 'html.parser')
        
        # Look for images
        images = soup.find_all('img')
        print(f"    🖼️  Found {len(images)} image(s)")
        
        downloaded_images = []
        for i, img in enumerate(images, 1):
            img_src = img.get('src', '')
            img_alt = img.get('alt', 'No alt')
            
            print(f"      Image {i}: {img_src[:80]}...")
            
            # Download the image if it's a Graph API URL
            if 'graph.microsoft.com' in img_src and '/onenote/resources/' in img_src:
                filename = download_image(img_src, i, safe_title)
                if filename:
                    downloaded_images.append({
                        'filename': filename,
                        'alt_text': img_alt,
                        'index': i
                    })
        
        # Check for object tags
        objects = soup.find_all('object')
        if objects:
            print(f"    📎 Found {len(objects)} object/embed(s)")
        
        # Extract text and create export
        text = soup.get_text(separator='\n', strip=True)
        
        # Create a markdown file with text and image references
        markdown_content = f"# {page['title']}\n\n"
        markdown_content += f"**Notebook:** {target_notebook['displayName']}\n"
        markdown_content += f"**Section:** {target_section['displayName']}\n"
        markdown_content += f"**Page:** {page_idx} of {len(all_pages)}\n\n"
        
        if text.strip():
            markdown_content += "## Content\n\n"
            markdown_content += text + "\n\n"
        
        if downloaded_images:
            markdown_content += "## Images\n\n"
            for img_info in downloaded_images:
                markdown_content += f"### Image {img_info['index']}\n"
                markdown_content += f"![Image {img_info['index']}](images/{img_info['filename']})\n\n"
                if img_info['alt_text'] and img_info['alt_text'] != 'No alt':
                    # Clean up the OCR text
                    clean_alt = img_info['alt_text'].replace('Machine generated alternative text:', '').strip()
                    if clean_alt:
                        markdown_content += f"**OCR Text from Image:**\n{clean_alt}\n\n"
        
        # Save the markdown file
        markdown_filename = f"{page_idx:02d}_{safe_title}.md"
        markdown_path = os.path.join(EXPORT_DIR, markdown_filename)
        
        with open(markdown_path, 'w', encoding='utf-8') as f:
            f.write(markdown_content)
        
        print(f"    📄 Exported to: {markdown_filename}")
        if downloaded_images:
            print(f"    🖼️  Downloaded {len(downloaded_images)} image(s)")
        
        # Show text preview
        if text.strip():
            preview = text[:200].replace('\n', ' ')
            print(f"    📝 Preview: {preview}..." if len(text) > 200 else f"    📝 Content: {preview}")
        else:
            print("    📝 (No text content)")

except requests.exceptions.RequestException as e:
    print(f"❌ API request failed: {e}")
    if hasattr(e, 'response') and e.response is not None:
        print(f"Response status: {e.response.status_code}")
        print(f"Response text: {e.response.text}")
except Exception as e:
    print(f"❌ An unexpected error occurred: {e}")

print(f"\n✅ Export completed! Check the '{EXPORT_DIR}' folder for all exported pages.")
