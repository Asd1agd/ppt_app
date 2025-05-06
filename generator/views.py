from django.shortcuts import render
from django.http import HttpResponse
from pptx import Presentation
from pptx.util import Pt
from io import BytesIO
import google.generativeai as genai
import sys
import io
import logging

# Setup logger
logger = logging.getLogger(__name__)
handler = logging.FileHandler('ppt_generation.log')  # logs stored here
formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
handler.setFormatter(formatter)
logger.addHandler(handler)
logger.setLevel(logging.INFO)


def home(request):
    if request.method == 'POST':
        log_stream = io.StringIO()
        sys.stdout = log_stream  # Redirect print

        try:
            api_key = request.POST['api_key']
            title = request.POST['title']

            # Configure Gemini API
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel('models/gemini-2.5-flash-preview-04-17')

            # Prompt for agenda
            prompt = f"give me presentation agenda list of 6 elements about: {title}, and give me the list separated by ',' and nothing else"
            response = model.generate_content(prompt)
            content = response.text.strip()
            print(f"Agenda generated: {content}")
            logger.info(f"Agenda: {content}")

            agenda_items = content.split(",")
            print(f"Agenda items: {agenda_items}")

            prs = Presentation()

            # Title slide
            slide = prs.slides.add_slide(prs.slide_layouts[0])
            slide.shapes.title.text = title
            slide.placeholders[1].text = "An AI-generated presentation"

            # Agenda slide
            slide = prs.slides.add_slide(prs.slide_layouts[1])
            slide.shapes.title.text = "Agenda"
            agenda_content = "\n".join([f"{i + 1}. {agenda.strip()}" for i, agenda in enumerate(agenda_items)])
            slide.placeholders[1].text = agenda_content

            def set_font_size(placeholder, font_size=18):
                if placeholder.has_text_frame:
                    for paragraph in placeholder.text_frame.paragraphs:
                        for run in paragraph.runs:
                            run.font.size = Pt(font_size)

            for i, agenda_item in enumerate(agenda_items):
                if agenda_item.strip():
                    slide = prs.slides.add_slide(prs.slide_layouts[1])
                    slide.shapes.title.text = f"Agenda {i + 1}: {agenda_item.strip()}"

                    agenda_content_response = model.generate_content(
                        f"Provide short content for the agenda item in 3 to 4 points: {agenda_item.strip()}"
                    )
                    agenda_content = agenda_content_response.text.strip().replace("*", "")

                    textbox = slide.shapes.placeholders[1]
                    textbox.text = agenda_content
                    set_font_size(textbox)

                    if len(agenda_content) > 800:
                        agenda_content = agenda_content[:800] + "..."
                        textbox.text = agenda_content
                        set_font_size(textbox)

                    print(f"Slide {i + 1} generated")

            ppt_io = BytesIO()
            prs.save(ppt_io)
            print("Presentation saved.")

            sys.stdout = sys.__stdout__  # Restore stdout
            logger.info("PPT generation complete.\n" + log_stream.getvalue())

            response = HttpResponse(
                ppt_io.getvalue(),
                content_type='application/vnd.openxmlformats-officedocument.presentationml.presentation'
            )
            response['Content-Disposition'] = 'attachment; filename=generated.pptx'

            return response

        except Exception as e:
            sys.stdout = sys.__stdout__
            logger.error("Error: " + str(e) + "\n" + log_stream.getvalue())
            return HttpResponse(f"<h3>Error occurred: {e}</h3><pre>{log_stream.getvalue()}</pre>")

    return render(request, 'form.html')
