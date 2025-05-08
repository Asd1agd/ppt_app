from django.shortcuts import render
from django.http import HttpResponse
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from io import BytesIO
import google.generativeai as genai
import random

# Load template
template = Presentation(rf"data/formatPpt2.pptx")
new_ppt = Presentation()
new_ppt.slide_width = template.slide_width
new_ppt.slide_height = template.slide_height

def Title(slide,data,font = 'Century Gothic',font_size =54,clr =[0,0,0], Top = 2.5):
    # Fixed left
    left = Inches(6.65)
    top = Inches(Top)
    font_size_pt = font_size
    sp = "                                                                                    \n"
    text = f"{sp}{data}"

    # Add textbox
    textbox = slide.shapes.add_textbox(left, top, 0, 0)
    text_frame = textbox.text_frame
    text_frame.text = text

    # Style
    p = text_frame.paragraphs[0]
    p.font.size = Pt(32)
    p.font.name = f'{font}'
    p.font.bold = True
    p.font.color.rgb = RGBColor(clr[0], clr[1], clr[2])
    p.alignment = PP_ALIGN.LEFT

    p = text_frame.paragraphs[1]
    p.font.size = Pt(font_size_pt)
    p.font.name = f'{font}'
    p.font.bold = True
    p.font.color.rgb = RGBColor(clr[0], clr[1], clr[2])
    p.alignment = PP_ALIGN.CENTER

def heading(slide,data,font = 'Century Gothic',font_size =32,clr =[0,0,0], Top = -0.3):
    # Fixed left
    left = Inches(6.65)
    top = Inches(Top)
    font_size_pt = font_size
    sp = "                                                                                    \n"
    text = f"{sp}  {data}"

    # Add textbox
    textbox = slide.shapes.add_textbox(left, top, 0, 0)
    text_frame = textbox.text_frame
    text_frame.text = text

    # Style
    p = text_frame.paragraphs[0]
    p.font.size = Pt(32)
    p.font.name = f'{font}'
    p.font.bold = True
    p.font.color.rgb = RGBColor(clr[0], clr[1], clr[2])
    p.alignment = PP_ALIGN.LEFT

    p = text_frame.paragraphs[1]
    p.font.size = Pt(font_size_pt)
    p.font.name = f'{font}'
    p.font.bold = True
    p.font.color.rgb = RGBColor(clr[0], clr[1], clr[2])
    p.alignment = PP_ALIGN.LEFT

    # textbox.fill.solid()
    # textbox.fill.fore_color.rgb = RGBColor(255, 255, 0)

def content(slide,data,font = 'Century Gothic',font_size=20,clr =[0,0,0], Top = 1,space=""):
    # Fixed left
    left = Inches(6.65)
    top = Inches(Top)
    font_size_pt = font_size
    sp = "                                                                                    \n"
    text = f"{sp}{data}"
    text = text.replace("\n",f"\n  {space}")

    # Add textbox
    textbox = slide.shapes.add_textbox(left, top, 0, 0)
    text_frame = textbox.text_frame
    text_frame.text = text

    # Style
    p = text_frame.paragraphs[0]
    p.font.size = Pt(32)
    p.font.name = f'{font}'
    p.font.bold = True
    p.font.color.rgb = RGBColor(clr[0], clr[1], clr[2])
    p.alignment = PP_ALIGN.LEFT

    for i in range(len(text.split('\n')) - 1):
        p = text_frame.paragraphs[i + 1]
        p.font.size = Pt(font_size_pt)
        p.font.name = f'{font}'
        p.font.bold = True
        p.font.color.rgb = RGBColor(clr[0], clr[1], clr[2])
        p.alignment = PP_ALIGN.LEFT

    # textbox.fill.solid()
    # textbox.fill.fore_color.rgb = RGBColor(255, 255, 0)


def home(request):
    if request.method == 'POST':
        api_key = request.POST['api_key']
        api_key = "AIzaSyBibUXKKGzdH-mErjMjNgYDY0kuxe4pK_I"
        title = request.POST['title']

        # Configure Gemini API
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel('models/gemini-2.5-flash-preview-04-17')

        # Prompt for the presentation agenda generation
        prompt = (
            f"give me presentation agenda list of 6 elements about: {title}, "
            f"and give me the list separated by ',' and nothing else. "
            f"No description, just 6 elements separated by commas."
        )
        try:
            # Generate content for the agenda
            response = model.generate_content(prompt)
            content = response.text.strip()
            print(f"[INFO] Agenda generated: {content}")
        except Exception as e:
            print(f"[ERROR] Gemini API failed: {e}")
            return HttpResponse(f"<h3>Error generating content: {e}</h3>")

        # Split the content into agenda items
        agenda_items = [item.strip() for item in content.split(",")]
        print(f"[INFO] Agenda items parsed: {agenda_items}")

        # Title Slide
        slide1 = new_ppt.slides.add_slide(template.slides[0].slide_layout)
        Title(slide1, prompt)

        # Agendas Slide
        slide2 = new_ppt.slides.add_slide(template.slides[10].slide_layout)
        heading(slide2, "               " + "Agenda")
        agenda_content = "\n".join([f"{i + 1}. {agenda}" for i, agenda in enumerate(agenda_items)])
        content(slide2, agenda_content, space="\n       ")

        # Individual agenda slides
        for i, item in enumerate(agenda_items):
            if item.strip():
                print(f"[INFO] Generating slide for agenda {i+1}: {item}")
                # Create a new slide with heading and content for each agenda item
                slide_no = random.randint(2, 6)
                slide = new_ppt.slides.add_slide(template.slides[slide_no].slide_layout)
                Agenda_no = f"Agenda {i + 1}: {item.strip()}"
                heading(slide, Agenda_no)

                try:
                    detail_response = model.generate_content(
                        f"Provide short content for the agenda item in 5 to 6 points: {item.strip()} and just give me the content points nothing else not any coment or description from your side"
                    )
                    agenda_details = detail_response.text.strip().replace("*", "")
                except Exception as e:
                    agenda_details = "Content generation failed."
                    print(f"[WARN] Failed content for {item}: {e}")

                # Add content to slide
                content(slide, agenda_details, space="\n    ")

        # Save and return presentation
        ppt_io = BytesIO()
        new_ppt.save(ppt_io)
        print("[INFO] PPT generation complete")

        response = HttpResponse(
            ppt_io.getvalue(),
            content_type='application/vnd.openxmlformats-officedocument.presentationml.presentation'
        )
        response['Content-Disposition'] = 'attachment; filename=generated.pptx'
        print("[INFO] PPT sent for download")
        return response

    return render(request, 'form.html')
