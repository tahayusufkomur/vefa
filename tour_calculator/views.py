# Create your views here.
from django.http import HttpResponse
from django.shortcuts import render


def index(request):
    if "GET" == request.method:
        return render(request, 'tour_calculator/index.html', {})

