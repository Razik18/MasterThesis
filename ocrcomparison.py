import os
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
import numpy as np
import pytesseract
import easyocr
from paddleocr import PaddleOCR
from google.cloud import vision
from azure.ai.vision.imageanalysis import ImageAnalysisClient
from azure.core.credentials import AzureKeyCredential
from Levenshtein import distance as levenshtein_distance



# Path to the folder containing the label images and ground truth .txt files
IMAGE_FOLDER_PATH = 'image_labels/'



#  Google Cloud Vision API:
os.environ['GOOGLE_APPLICATION_CREDENTIALS'] = ''

#Azure AI Vision API:
AZURE_ENDPOINT = ''
AZURE_KEY = ''



def calculate_cer(ground_truth, ocr_output):
    """Calculates the Character Error Rate (CER) between two strings."""
    if ground_truth is None or ocr_output is None:
        return 1.0 # Max error if one is None
    gt_clean = ' '.join(ground_truth.strip().split())
    ocr_clean = ' '.join(ocr_output.strip().split())
    if len(gt_clean) == 0:
        return 0.0 if len(ocr_clean) == 0 else 1.0
    
    lev_dist = levenshtein_distance(gt_clean, ocr_clean)
    cer = lev_dist / len(gt_clean)
    return cer



reader_easyocr = easyocr.Reader(['en'])
ocr_paddle = PaddleOCR(use_angle_cls=True, lang='en')
client_google = vision.ImageAnnotatorClient()
client_azure = ImageAnalysisClient(endpoint=AZURE_ENDPOINT, credential=AzureKeyCredential(AZURE_KEY))

def run_tesseract(image_path):
    try:
        return pytesseract.image_to_string(image_path)
    except Exception as e:
        print(f"Tesseract Error: {e}")
        return ""

def run_easyocr(image_path):
    try:
        result = reader_easyocr.readtext(image_path, detail=0, paragraph=True)
        return ' '.join(result)
    except Exception as e:
        print(f"EasyOCR Error: {e}")
        return ""

def run_paddleocr(image_path):
    try:
        result = ocr_paddle.ocr(image_path, cls=True)
        if result and result[0]:
            lines = [line[1][0] for line in result[0]]
            return ' '.join(lines)
        return ""
    except Exception as e:
        print(f"PaddleOCR Error: {e}")
        return ""

def run_google_vision(image_path):
    try:
        with open(image_path, 'rb') as image_file:
            content = image_file.read()
        image = vision.Image(content=content)
        response = client_google.text_detection(image=image)
        if response.text_annotations:
            return response.text_annotations[0].description
        return ""
    except Exception as e:
        print(f"Google Vision Error: {e}")
        return ""

def run_azure_vision(image_path):
    try:
        with open(image_path, "rb") as f:
            image_data = f.read()
        
        result = client_azure.analyze(
            image_data=image_data,
            visual_features=['read']
        )
        if result.read and result.read.blocks:
            return result.read.content
        return ""
    except Exception as e:
        print(f"Azure Vision Error: {e}")
        return ""


def run_benchmark():
    """Runs the full OCR benchmark on all images in the folder."""
    results = []
    image_files = sorted([f for f in os.listdir(IMAGE_FOLDER_PATH) if f.lower().endswith(('.png', '.jpg', '.jpeg'))])

    if not image_files:
        print(f"Error: No image files found in '{IMAGE_FOLDER_PATH}'")
        return None

    for i, filename in enumerate(image_files):
        print(f"Processing image {i+1}/{len(image_files)}: {filename}")
        image_path = os.path.join(IMAGE_FOLDER_PATH, filename)
        base_name = os.path.splitext(filename)[0]
        ground_truth_path = os.path.join(IMAGE_FOLDER_PATH, f"{base_name}.txt")

        if not os.path.exists(ground_truth_path):
            print(f"  - Warning: Ground truth file not found for {filename}. Skipping.")
            continue

        with open(ground_truth_path, 'r', encoding='utf-8') as f:
            ground_truth = f.read()

        # Run each OCR engine
        ocr_outputs = {
            "Tesseract": run_tesseract(image_path),
            "EasyOCR": run_easyocr(image_path),
            "PaddleOCR": run_paddleocr(image_path),
            "Google Vision": run_google_vision(image_path),
            "Azure Vision": run_azure_vision(image_path)
        }
        
        # Calculate CER for each engine
        for engine, output in ocr_outputs.items():
            cer = calculate_cer(ground_truth, output)
            results.append({
                'Image': filename,
                'OCR Engine': engine,
                'Character Error Rate (%)': cer * 100
            })

    return pd.DataFrame(results)

# --- Run the script and generate plots ---
if __name__ == "__main__":
    results_df = run_benchmark()

    if results_df is not None and not results_df.empty:
        sns.set_theme(style="whitegrid")
        plt.figure(figsize=(12, 8))
        ax = sns.boxplot(
            x='OCR Engine', 
            y='Character Error Rate (%)', 
            data=results_df,
            palette="Set2",
            order=["Tesseract", "EasyOCR", "PaddleOCR", "Azure Vision", "Google Vision"]
        )
        plt.title('Comparison of OCR Engine Performance on Slide Labels (N={})'.format(len(results_df['Image'].unique())), fontsize=16, pad=20)
        plt.xlabel('OCR Engine', fontsize=12)
        plt.ylabel('Character Error Rate (CER %)', fontsize=12)
        plt.tight_layout()
        plt.savefig("ocr_benchmark_boxplot_from_script.png", dpi=300)
        print("\nBox plot saved as 'ocr_benchmark_boxplot_from_script.png'")


        pivot_df = results_df.pivot(index='Image', columns='OCR Engine', values='Character Error Rate (%)')
        print("\nDescriptive Statistics for OCR Character Error Rate (%):")
        print(pivot_df.describe().round(2))
    else:
        print("\nBenchmark could not be completed.")