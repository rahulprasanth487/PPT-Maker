import fitz  # Import PyMuPDF

def extract_images_from_pdf(pdf_path):
    """Extracts images from a PDF and saves them with sequential names.

    Args:   
        pdf_path (str): The path to the PDF file.
    """

    doc = fitz.open(pdf_path)  # Open the PDF

    image_count = 1
    for page_num in range(len(doc)):
        page = doc[page_num]
        images = page.get_images(full=True)  # Get all image objects on the page

        for img_index, img in enumerate(images):
            xref = img[0]
            base_image = doc.extract_image(xref)

            image_data = base_image["image"]
            image_ext = base_image["ext"]

            image_filename = f"images/image{image_count}.{image_ext}"
            with open(image_filename, "wb") as image_file:
                image_file.write(image_data)
            print(f"Saved image: {image_filename}")
            image_count += 1

# --- Example Usage ---
pdf_file_path = "slice.pdf"  # Replace with your PDF's path
extract_images_from_pdf(pdf_file_path) 