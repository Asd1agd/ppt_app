from django.shortcuts import render
from django.http import HttpResponse
from pptx import Presentation
from io import BytesIO
import google.generativeai as genai

def home(request):
    if request.method == 'POST':
        api_key = request.POST['api_key']
        title = request.POST['title']

        # Configure Gemini API
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel('gemini-pro')
        prompt = f"Write a short presentation content (one paragraph) about: {title}"
        try:
            response = model.generate_content(prompt)
            content = response.text.strip()
        except Exception as e:
            return HttpResponse(f"<h3>Error generating content: {e}</h3>")

        # Create PPT
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        slide.shapes.title.text = title
        slide.placeholders[1].text = content

        ppt_io = BytesIO()
        prs.save(ppt_io)
        ppt_io.seek(0)

        response = HttpResponse(
            ppt_io.read(),
            content_type='application/vnd.openxmlformats-officedocument.presentationml.presentation'
        )
        response['Content-Disposition'] = 'attachment; filename=generated.pptx'
        return response

    return render(request, 'form.html')
