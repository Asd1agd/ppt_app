# AIzaSyBibUXKKGzdH-mErjMjNgYDY0kuxe4pK_I

import google.generativeai as genai
from pptx import Presentation
from pptx.util import Pt
from io import BytesIO

# Configure Gemini
api_key = input("Enter your Gemini API key: ")
genai.configure(api_key="AIzaSyBibUXKKGzdH-mErjMjNgYDY0kuxe4pK_I")

model = genai.GenerativeModel('models/gemini-2.5-flash-preview-04-17')
prompt = input("Enter the PPT title: ")

# Generate content for the presentation
try:
    response = model.generate_content(f"give me presentation agenda list of 6 elements about: {prompt} , and give me the list separated by ',' and noting else not description or else just 10 elements separated by ','")
    content = response.text.strip()
except Exception as e:
    print(f"Error generating content: {e}")
    exit()

# Split the content into agenda items
agenda_items = content.split(",")
print(agenda_items)

# Create a PowerPoint presentation
prs = Presentation()

# Title Slide
slide = prs.slides.add_slide(prs.slide_layouts[0])
slide.shapes.title.text = prompt
slide.placeholders[1].text = "An AI-generated presentation"

# Agendas Slide
slide = prs.slides.add_slide(prs.slide_layouts[1])
slide.shapes.title.text = "Agenda"
agenda_content = "\n".join([f"{i+1}. {agenda}" for i, agenda in enumerate(agenda_items)])
slide.placeholders[1].text = agenda_content

# Helper function to set font size and fit text
def set_font_size(placeholder, font_size=18):
    # Accessing the text frame directly
    if placeholder.has_text_frame:
        for paragraph in placeholder.text_frame.paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(font_size)

# Create individual slides for each agenda item
for i, agenda_item in enumerate(agenda_items):
    if agenda_item.strip():  # Skip any empty agenda item
        # Create a new slide with heading and content for each agenda item
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = f"Agenda {i+1}: {agenda_item.strip()}"

        # Generate content for each agenda item (optional, adjust logic as needed)
        agenda_content_response = model.generate_content(f"Provide short content for the agenda item in 3 to 4 points: {agenda_item.strip()}")
        agenda_content = agenda_content_response.text.strip()

        # Remove any unwanted formatting (like stars for bold)
        agenda_content = agenda_content.replace("*", "")  # Remove stars used for bold formatting

        # Add the content related to the agenda item
        textbox = slide.shapes.placeholders[1]
        textbox.text = agenda_content

        # Set font size for the content
        set_font_size(textbox, font_size=18)

        # Ensure content fits within the slide
        if len(agenda_content) > 800:  # If content is too long, trim it
            agenda_content = agenda_content[:800] + "..."
            textbox.text = agenda_content
            set_font_size(textbox, font_size=18)

        print(f"Slide {i+1} generated")

# Save the presentation to a file
ppt_io = BytesIO()
prs.save(ppt_io)

with open("generated.pptx", "wb") as f:
    f.write(ppt_io.getbuffer())

print("Presentation generated: generated.pptx")
