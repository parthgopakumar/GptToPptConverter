from pptx import Presentation
from pptx.util import Pt
import re

def create_presentation(input_file, output_file):
    # Read the contents of the input file
    with open(input_file, 'r') as file:
        slides_content = file.read()
    
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

        # Organization of the presentation
        for line in lines[1:]:
            if line.strip():  # Avoid empty lines
                # Add a new paragraph for each line of content
                p = content.add_paragraph()
                p.text = line.strip()
                p.space_after = Pt(14)  # Add some space after each paragraph

    # Save the presentation
    prs.save(output_file)

# Specify the input and output files
input_file = 'slides_input.txt'  # Input text file
output_file = 'output.pptx' # Output PPT file 

# Create the PowerPoint presentation
create_presentation(input_file, output_file)

print(f"Presentation saved as '{output_file}'")
