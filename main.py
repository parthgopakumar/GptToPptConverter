#imports required 
from pptx import Presentation
from pptx.util import Pt
import re

def create_presentation(slides_content, output_file):
    # Regular expression to split the slides
    slides = re.split(r'\n(?=Slide \d+:)', slides_content.strip())

    # Create a presentation object
    prs = Presentation()

    for slide_text in slides:
        lines = slide_text.split('\n')
        # First line is the title
        slide_title = lines[0].split(': ')[1].strip()
        slide = prs.slides.add_slide(prs.slide_layouts[1])  # Adding a title and content slide
        title = slide.shapes.title
        title.text = slide_title
        
        # Rest of the lines are content, add them to the content placeholder
        content = slide.placeholders[1].text_frame
        content.text = ""  # Initialize text frame

        for line in lines[1:]:
            if line.strip():  # Avoid empty lines
                # Add a new paragraph for each line of content
                p = content.add_paragraph()
                p.text = line.strip()
                p.space_after = Pt(14)  # Add some space after each paragraph

    # Save the presentation
    prs.save(output_file)

# Function to get user input
def get_slides_input():
    print("Use the format 'Slide n: Title' for slide titles.")
    print("Enter the slide content line by line. Enter an empty line when you're done with a slide.")
    print("Type 'end' on a new line when you are done entering all slides. (This should be at the very end) \n")

    slides_content = ""
    while True:
        line = input()
        if line.strip().upper() == "end":
            break
        slides_content += line + "\n"

    return slides_content

# Get user input for slides
slides_content = get_slides_input()

# Specify the output file (check files/folders for this specific file)
output_file = 'AI_Introduction_Presentation.pptx'

# Create the PowerPoint presentation
create_presentation(slides_content, output_file)

print(f"Presentation saved as '{output_file}'")
