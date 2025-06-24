from pptx import Presentation
from transformers import pipeline
import os
import re

def extract_text_and_images_from_pptx(pptx_file, output_image_dir):
    prs = Presentation(pptx_file)
    slides_content = []
    
    os.makedirs(output_image_dir, exist_ok=True)

    for i, slide in enumerate(prs.slides):
        slide_content = {"text": "", "images": []}
        
        title = ""
        for shape in slide.shapes:
            if hasattr(shape, "text") and shape.text.strip():
                title = shape.text.strip().split("\n")[0]  
                break

        slide_content["title"] = title if title else f"Slide {i+1}"
        
        for shape in slide.shapes:
            if hasattr(shape, "text"):
                text = shape.text.strip().replace("�", "")  
                text = re.sub(r'[^\x00-\x7F]+', '', text) 
                
                bullet_points = re.split(r'(?<!\w\.\w.)(?<![A-Z][a-z]\.)(?<=\.|\?)\s', text)
                slide_content["text"] += "\n".join([f"- {point.strip()}" for point in bullet_points if point.strip()]) + "\n"

        for shape in slide.shapes:
            if hasattr(shape, "image"):
                image = shape.image
                image_filename = os.path.join(output_image_dir, f"slide_{i+1}_image_{len(slide_content['images'])+1}.png")
                with open(image_filename, "wb") as img_file:
                    img_file.write(image.blob)
                slide_content["images"].append(image_filename)

        slides_content.append(slide_content)

    return slides_content

def summarise_slides(slides_content):
    summariser = pipeline("summarization", model="facebook/bart-large-cnn")
    summarised_notes = []
    
    for slide in slides_content:
        text = slide["text"].strip()
        if len(text) > 0: 
            summary = summariser(text, max_length=250, min_length=80, do_sample=False)[0]['summary_text']
            summary_bullets = re.split(r'(?<!\w\.\w.)(?<![A-Z][a-z]\.)(?<=\.|\?)\s', summary)
            bullet_summary = "\n".join([f"- {s.strip()}" for s in summary_bullets if s.strip()])
        else:
            bullet_summary = "(No significant text on this slide)"
        summarised_notes.append({"title": slide["title"], "summary": bullet_summary, "images": slide["images"]})
    
    return summarised_notes

def save_as_markdown(summarised_notes, output_file, image_dir):
    with open(output_file, "w") as f:
        for i, slide in enumerate(summarised_notes):
            f.write(f"# {slide['title']}\n") 
            f.write(f"{slide['summary']}\n\n")
            
            for img_path in slide["images"]:
                relative_path = os.path.relpath(img_path, image_dir)
                f.write(f'<img src="{relative_path}" width="500px">\n')
            f.write("\n")

def process_pptx_to_detailed_notes(pptx_file, output_file, output_image_dir):
    slides_content = extract_text_and_images_from_pptx(pptx_file, output_image_dir)
    summarised_notes = summarise_slides(slides_content)
    save_as_markdown(summarised_notes, output_file, output_image_dir)
    print(f"summarised notes with images saved to {output_file}")

pptx_file = "W1L1-comp-org-intro.pptx" 
output_file = "notes_for_obsidian.md"
output_image_dir = r"C:\Users\" 

process_pptx_to_detailed_notes(pptx_file, output_file, output_image_dir)
