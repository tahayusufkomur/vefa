from django.http import HttpResponse
from django.shortcuts import render
from django.contrib.auth.decorators import login_required

@login_required
def index(response):
    return render(response, "vefa/base.html", {})

@login_required
def home(response):
    return render(response, "vefa/home.html", {})