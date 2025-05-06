from django.shortcuts import render
from django.http import HttpResponse
from pptx import Presentation
from io import BytesIO

def home(request):
    if request.method == 'POST':
        title = request.POST['title']
        content = request.POST['content']

        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        slide.shapes.title.text = title
        slide.placeholders[1].text = content

        ppt_io = BytesIO()
        prs.save(ppt_io)
        ppt_io.seek(0)

        response = HttpResponse(ppt_io.read(), content_type='application/vnd.openxmlformats-officedocument.presentationml.presentation')
        response['Content-Disposition'] = 'attachment; filename=generated.pptx'
        return response

    return render(request, 'form.html')
