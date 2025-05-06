from django.shortcuts import render
from django.http import HttpResponse
from pptx import Presentation
from pptx.util import Pt
from io import BytesIO
import google.generativeai as genai

def home(request):
    if request.method == 'POST':
        api_key = request.POST['api_key']
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

        # Create a PowerPoint presentation
        prs = Presentation()

        # Title Slide
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        slide.shapes.title.text = title
        slide.placeholders[1].text = "An AI-generated presentation"

        # Agenda Slide
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = "Agenda"
        agenda_text = "\n".join([f"{i+1}. {agenda}" for i, agenda in enumerate(agenda_items)])
        slide.placeholders[1].text = agenda_text

        # Set font size helper
        def set_font_size(placeholder, font_size=18):
            if placeholder.has_text_frame:
                for paragraph in placeholder.text_frame.paragraphs:
                    for run in paragraph.runs:
                        run.font.size = Pt(font_size)

        # Individual agenda slides
        for i, item in enumerate(agenda_items):
            if item:
                print(f"[INFO] Generating slide for agenda {i+1}: {item}")
                slide = prs.slides.add_slide(prs.slide_layouts[1])
                slide.shapes.title.text = f"Agenda {i+1}: {item}"

                try:
                    detail_response = model.generate_content(
                        f"Provide short content in 3 to 4 points for: {item}"
                    )
                    agenda_details = detail_response.text.strip().replace("*", "")
                except Exception as e:
                    agenda_details = "Content generation failed."
                    print(f"[WARN] Failed content for {item}: {e}")

                # Add content to slide
                textbox = slide.placeholders[1]
                if len(agenda_details) > 800:
                    agenda_details = agenda_details[:800] + "..."
                textbox.text = agenda_details
                set_font_size(textbox)

        # Save and return presentation
        ppt_io = BytesIO()
        prs.save(ppt_io)
        print("[INFO] PPT generation complete")

        response = HttpResponse(
            ppt_io.getvalue(),
            content_type='application/vnd.openxmlformats-officedocument.presentationml.presentation'
        )
        response['Content-Disposition'] = 'attachment; filename=generated.pptx'
        print("[INFO] PPT sent for download")
        return response

    return render(request, 'form.html')
