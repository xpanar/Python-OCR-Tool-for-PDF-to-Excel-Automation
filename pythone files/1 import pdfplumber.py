import pdfplumber
with pdfplumber.open("C:/Users/xpana/OneDrive/Desktop/big qoutation.pdf") as pdf:
    # Extract the text
    text = pdf.extract_text()
    print(text)

    # Extract the data
    tables = pdf.extract_table()
    for table in tables:
        print(table)

    # Extract the images
    images = pdf.get_images()
    for image in images:
        print(image["page_number"])
        with open(f"image_{image['page_number']}.jpg", "wb") as f:
            f.write(image["data"])