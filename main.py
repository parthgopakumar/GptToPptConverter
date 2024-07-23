import os
from pptx import Presentation
from pptx.util import Inches, Pt

def create_slide(prs, title, content):
    slide_layout = prs.slide_layouts[1]  # Using a slide with title and content
    slide = prs.slides.add_slide(slide_layout)
    
    # Set the slide title
    title_shape = slide.shapes.title
    title_shape.text = title
    
    # Set the slide content
    content_shape = slide.placeholders[1]
    tf = content_shape.text_frame
    tf.text = content

def main():
    # Get user input
    presentation_title = input("Enter your presentation title: ")
    
    # creates a new presentation
    prs = Presentation()
    
    # Create title slide
    slide_layout = prs.slide_layouts[0]  
    slide = prs.slides.add_slide(slide_layout)
    title_shape = slide.shapes.title
    title_shape.text = presentation_title
    
    # Get content for slides
    slides = []
    while True:
        slide_title = input("Enter slide title (or press Enter to finish): ")
        if not slide_title:
            break
        
        print("Enter slide content (press Enter twice to finish):")
        lines = []
        while True:
            line = input()
            if line:
                lines.append(line)
            else:
                break
        slide_content = "\n".join(lines)
        
        slides.append((slide_title, slide_content))
    
    # Create content slides
    for title, content in slides:
        create_slide(prs, title, content)
    
    # Save the presentation
    prs.save('generated_presentation.pptx')
    print(f"Presentation saved as 'generated_presentation.pptx' in {os.getcwd()}")

if __name__ == "__main__":
    main()


# How to Run
# pip3 install python-pptx
# run code and enter info
