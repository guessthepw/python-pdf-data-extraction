import imghdr
import os
from flask import Flask, render_template, request, redirect, url_for, abort, \
    send_from_directory
from werkzeug.utils import secure_filename
import pdf2image
try:
    from PIL import Image
except ImportError:
    import Image
import pytesseract
import fitz # PyMuPDF
import io
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Inches
import sys
import os

tessdata_dir_config = '--tessdata-dir "/Users/johnherbener/tesseract/tessdata"'
doc = DocxTemplate("template.docx")
specs = ""
ph1 = InlineImage(doc, image_descriptor="p1.png") #Lockson Logo
outside_context = { "context" : [] }
app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 1024 * 1024
app.config['UPLOAD_EXTENSIONS'] = ['.jpg', '.png', '.gif', '.pdf']
app.config['UPLOAD_PATH'] = 'uploads'
global log
log = ' '

def validate_image(stream):
    header = stream.read(512)  # 512 bytes should be enough for a header check
    stream.seek(0)  # reset stream pointer
    format = imghdr.what(None, header)
    if not format:
        return None
    return '.' + (format if format != 'jpeg' else 'jpg')

def pdf_to_img(pdf):
    print("[+] Converting PDF to JPEG", file=sys.stderr)
    global log 
    log += "[+] Converting PDF to JPEG"
    return pdf2image.convert_from_path(pdf)

def ocr_core(file):
    print("[+] Reading Text from the JPEG")
    global log 
    log += "[+] Reading Text from the JPEG"
    text = pytesseract.image_to_string(file, config=tessdata_dir_config)
    return text

def print_pages(pdf):
    print("[+] Opening PDF")
    global log 
    log += "[+] Opening PDF"
    pdf_file = fitz.open(pdf)
    images = pdf_to_img(pdf)
    myimages = [
        { "img" : ph1 }
    ]
    print("[+] Starting Data Extraction")
    log += "[+] Starting Data Extraction"
    for pg, img in enumerate(images):
        print("[+] Starting Text Extraction")
        log += "[+] Starting Text Extraction"
        specs = ocr_core(img)
        specs = specs.replace('\n\n', ' ')
        print("[+] Finished Text Extraction")
        log += "[+] Finished Text Extraction"
        # get the page itself
        page = pdf_file[pg]
        image_list = page.getImageList()
        print("[+] Starting Image Extraction")
        # printing number of images found in this page
        if image_list:
            print(f"[+] Found a Total of {len(image_list)} Images in Page {pg}")
            log += f"[+] Found a Total of {len(image_list)} Images in Page {pg}"
        else:
            print("[!] No images found on page", pg)
        print("[+] Saving Images")
        log += "[+] Saving Images"
        for image_index, img in enumerate(page.getImageList(), start=1):
            # get the XREF of the image
            xref = img[0]
            # extract the image bytes
            base_image = pdf_file.extractImage(xref)
            image_bytes = base_image["image"]
            # get the image extension
            image_ext = base_image["ext"]
            # load it to PIL
            image = Image.open(io.BytesIO(image_bytes))
            # save it to local disk
            fnpg = pg+1
            fn = "images/image" + str(fnpg) + "_" + str(image_index) + "." + str(image_ext)
            image.save(open(f"images/image{pg+1}_{image_index}.{image_ext}", "wb"))
            myimages.append({ "img" : InlineImage(doc, image_descriptor=fn, width=Inches(2.5))})    
            print(f"[+] Finished Image Extraction {image_index}")
            log += "[+] Finished Image Extraction"
        context = { 'string' : specs,
                    'image' : myimages
        }
        outside_context.get("context").append(context)
        specs = ""
        myimages = [
            { "img" : ph1 }
        ]
        print(f"[+] Saving Data Page {pg}")
        log += "[+] Saving Page Data"

    print("[+] Saving Document")
    log += "[+] Saving Document"
    doc.render(outside_context)
    size = len(pdf)
    doc.save(f"{pdf[:size - 4]}.docx")


@app.route('/')
def index():
    files = os.listdir(app.config['UPLOAD_PATH'])
    global log
    return render_template('index.html', files=files)

@app.route('/', methods=['POST'])
def upload_files():
    uploaded_file = request.files['file']
    filename = secure_filename(uploaded_file.filename)
    if filename != '':
        file_ext = os.path.splitext(filename)[1]
        if file_ext not in app.config['UPLOAD_EXTENSIONS']:
            abort(400)
        uploaded_file.save(os.path.join(app.config['UPLOAD_PATH'], filename))
        print_pages(f"uploads/{filename}")
    return redirect(f"/uploads/{filename[:len(filename) - 4]}.docx")

@app.route('/uploads/<filename>')
def upload(filename):
    return send_from_directory(app.config['UPLOAD_PATH'], filename)


if __name__ == '__main__':

    # Start app
    app.run()
