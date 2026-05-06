import cv2
import numpy as np
import base64
import sys
import os
import json
from openai import OpenAI
from PIL import Image
import io

client = OpenAI()

def crop_receipt(image_path):
    # Carregar imagem
    img = cv2.imread(image_path)
    if img is None:
        return None, "Não foi possível carregar a imagem."
    
    # Converter para escala de cinza e aplicar blur
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    blur = cv2.GaussianBlur(gray, (5, 5), 0)
    
    # Detecção de bordas
    edged = cv2.Canny(blur, 75, 200)
    
    # Encontrar contornos
    cnts, _ = cv2.findContours(edged.copy(), cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    cnts = sorted(cnts, key=cv2.contourArea, reverse=True)
    
    receipt_cnt = None
    for c in cnts:
        peri = cv2.arcLength(c, True)
        approx = cv2.approxPolyDP(c, 0.02 * peri, True)
        if len(approx) == 4:
            receipt_cnt = approx
            break
            
    if receipt_cnt is not None:
        # Transformação de perspectiva (simplificada para recorte retangular)
        x, y, w, h = cv2.boundingRect(receipt_cnt)
        cropped = img[y:y+h, x:x+w]
    else:
        # Se não encontrar contorno de 4 pontos, apenas tenta remover as bordas vazias
        # ou retorna a imagem original se falhar
        cropped = img
        
    output_path = "processed_" + os.path.basename(image_path)
    cv2.imwrite(output_path, cropped)
    return output_path, None

def extract_data_with_llm(image_path):
    with open(image_path, "rb") as image_file:
        base64_image = base64.b64encode(image_file.read()).decode('utf-8')

    response = client.chat.completions.create(
        model="gpt-4.1-mini",
        messages=[
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": "Analise este recibo e extraia as seguintes informações em formato JSON: concessionaria, data (DD/MM/AAAA), hora, placa, valor (apenas número), tipo (pedagio ou estacionamento). Se for pedágio, identifique a concessionária. Se for estacionamento, identifique o local."},
                    {
                        "type": "image_url",
                        "image_url": {
                            "url": f"data:image/jpeg;base64,{base64_image}",
                        },
                    },
                ],
            }
        ],
        response_format={ "type": "json_object" }
    )
    return response.choices[0].message.content

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(json.dumps({"error": "Nenhuma imagem fornecida."}))
        sys.exit(1)
        
    img_path = sys.argv[1]
    
    try:
        # 1. Recortar imagem
        processed_path, error = crop_receipt(img_path)
        if error:
            print(json.dumps({"error": error}))
            sys.exit(1)
            
        # 2. OCR e extração de dados
        data_json_str = extract_data_with_llm(processed_path)
        data = json.loads(data_json_str)
        
        # 3. Adicionar o caminho da imagem processada
        data["processed_image"] = processed_path
        
        print(json.dumps(data))
        
    except Exception as e:
        print(json.dumps({"error": str(e)}))
        sys.exit(1)
