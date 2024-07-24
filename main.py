from pptx import Presentation
from pptx.util import Pt
import re

def new_presentation(input_file, output_file):
    # Read the contents of the input file
    with open(input_file, 'r') as file:
        slides_content = file.read()
    
    # Regular Expression Explanation
    
    # Regular expression to split the slides
    # slides_content.strip() removes trailing whitepsace from the slides_content string, so that it does not interfere with splitting
    # re.split() splits the string wherever the regular expression matches with a pattern
    # r represents raw string literal
    # \n is newline character and this is where splitting occurs
    # The info in the paranthesis is the lookahead assertion, which is match the input with the given pattern
    # In this context, we define the split where it is after string with Slide and d+ or digit, and it will confirm if this is true before spliting
    
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
new_presentation(input_file, output_file)

print(f"Presentation saved as '{output_file}'")
