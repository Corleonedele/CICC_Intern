from django.urls import path
from django.conf.urls.static import static
from django.conf import settings
from . import views

urlpatterns = [
    path('', views.index, name='index'),
    path('追加',views.追加, name='追加'),
    path('分红再投',views.分红再投, name='分红再投'),
    path('现金分红',views.现金分红, name='现金分红'),
    path('赎回',views.赎回, name='赎回'),
    path('调减',views.调减, name='调减'),
    path('IC查询',views.IC查询, name='IC查询'),

    
    path('download', views.下载日报, name="下载日报"),
    path('test',views.test, name='test'),
] + static(settings.STATIC_URL, document_root=settings.STATIC_ROOT)