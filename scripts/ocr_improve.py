import sys
import os
from pathlib import Path
import cv2
import numpy as np
from PIL import Image
import pytesseract
import easyocr

reader = None


def preprocess_image(img_path, scale=2):
    img = cv2.imread(str(img_path))
    if img is None:
        return None
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    # upscale to improve OCR
    h, w = gray.shape
    gray = cv2.resize(gray, (w*scale, h*scale), interpolation=cv2.INTER_CUBIC)
    # median blur to reduce noise
    gray = cv2.medianBlur(gray, 3)
    # adaptive threshold
    th = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                               cv2.THRESH_BINARY, 31, 12)
    return th


def ocr_images(images_dir, out_txt):
    images = sorted([p for p in Path(images_dir).iterdir() if p.suffix.lower() in ['.png', '.jpg', '.jpeg', '.tiff']])
    if not images:
        print('No images found in', images_dir)
        return 1

    results = []
    global reader
    if reader is None:
        try:
            reader = easyocr.Reader(['es'], gpu=False)
        except Exception as e:
            print('Failed to initialize EasyOCR reader:', e)
            return 3

    for img_path in images:
        print('Processing', img_path.name)
        pre = preprocess_image(img_path)
        if pre is None:
            print('  Skipped (could not read)')
            continue
        # save temp for debug
        tmp_path = Path(images_dir) / f'pre_{img_path.name}'
        cv2.imwrite(str(tmp_path), pre)
        text = ''
        try:
            pil = Image.fromarray(pre)
            text = pytesseract.image_to_string(pil, lang='spa')
        except pytesseract.pytesseract.TesseractNotFoundError:
            print('Tesseract binary not found, falling back to EasyOCR...')
            try:
                results_ocr = reader.readtext(np.array(pre), detail=0, paragraph=True)
                text = '\n'.join(results_ocr)
            except Exception as e:
                print('  EasyOCR error:', e)
                text = ''
        except Exception as e:
            print('  OCR error (pytesseract):', e)
            try:
                results_ocr = reader.readtext(np.array(pre), detail=0, paragraph=True)
                text = '\n'.join(results_ocr)
            except Exception as e2:
                print('  EasyOCR error:', e2)
                text = ''
        results.append((img_path.name, text))

    with open(out_txt, 'w', encoding='utf-8') as f:
        for name, text in results:
            f.write(f'# Image: {name}\n\n')
            f.write(text.strip() + '\n\n')

    print('OCR finished. Output saved to', out_txt)
    return 0


if __name__ == '__main__':
    if len(sys.argv) > 1:
        images_dir = sys.argv[1]
    else:
        images_dir = 'output/20260528_093910_6240/images'
    out_txt = Path(images_dir).parent / 'ocr_improved.txt'
    sys.exit(ocr_images(images_dir, out_txt))
