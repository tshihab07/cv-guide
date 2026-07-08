from django.urls import path

from . import views

urlpatterns = [
    path('', views.landing_page, name='landing_page'),
    path('create/', views.builder, name='builder'),
    path('preview/', views.preview, name='preview'),
    path('download/pdf/', views.download_pdf, name='download_pdf'),
    path('download/docx/', views.download_docx, name='download_docx'),
    path('cv-admin/', views.cv_admin_dashboard, name='cv_admin_dashboard'),
    path('cv-admin/login/', views.cv_admin_login, name='cv_admin_login'),
    path('cv-admin/logout/', views.cv_admin_logout, name='cv_admin_logout'),
]
