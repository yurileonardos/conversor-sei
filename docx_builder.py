def process_tr(pdf_bytes):
    for page in pdf_pages:
        image = pdf_page_to_image()
        if contains_table(image):
            image = apply_visual_mask(image)
        yield image


def process_proposal(pdf_bytes):
    for page in pdf_pages:
        image = pdf_page_to_image()
        yield image
