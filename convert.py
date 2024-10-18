from docx import Document
import os

# Paths
doc_path = '/Users/dma/Downloads/_Cortex XDR Event Forwarding to Google SecOps (Chronicle).docx'  # Replace with your actual path
img_folder = './images'  # Path where images will be stored

# Create a folder to store images if it doesn't exist
os.makedirs(img_folder, exist_ok=True)

# Load the DOCX file
doc = Document(doc_path)

# Prepare to collect text and image references
markdown_content_with_images = []

# Extract text and images
for para in doc.paragraphs:
    markdown_content_with_images.append(para.text)  # Add paragraph text
    for run in para.runs:
        # Check if the run has an inline shape (image)
        for shape in run._element.findall('.//a:blip', namespaces={'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}):
            # Get the corresponding relationship id
            rel_id = shape.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
            if rel_id:
                # Retrieve the image part using the relationship id
                image_part = doc.part.related_parts[rel_id]
                img_filename = os.path.basename(image_part.partname)
                img_path = os.path.join(img_folder, img_filename)

                # Save the image to the specified folder
                with open(img_path, 'wb') as img_file:
                    img_file.write(image_part.blob)
                
                # Add the Markdown image reference
                markdown_content_with_images.append(f"![Image](./{os.path.basename(img_folder)}/{img_filename})")

# Join the content into a single string
final_markdown_content = "\n\n".join(markdown_content_with_images)

# Save the updated markdown with images in the correct positions
updated_md_file = './updated_markdown_file_with_images.md'  # Replace with the desired output path
with open(updated_md_file, 'w') as md_file:
    md_file.write(final_markdown_content)

print(f"Markdown file with images saved as: {updated_md_file}")