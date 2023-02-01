#-*-coding:utf-8-*-
from django.shortcuts import render
from django.http import HttpResponse
from django.template import loader
import json
# Create your views here.



def index(request):
    return render(request, 'root/index.html')
    # return HttpResponse(template.render(context, request))


def test(request):
    if request.method == 'POST':
        # body_dict = (request.body.decode('utf-8'))
        body_dict = request.POST.dict()
        print(body_dict)
    return HttpResponse(200)





def 追加(request):
    return render(request, 'op/追加.html')

def 调减(request):
    return render(request, 'op/调减.html')

def 赎回(request):
    return render(request, 'op/赎回.html')

def 现金分红(request):
    return render(request, 'op/现金分红.html')

def 分红再投(request):
    return render(request, 'op/分红再投.html')
