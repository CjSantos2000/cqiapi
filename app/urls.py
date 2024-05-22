from django.urls import path
from . import views

urlpatterns = [
    path('obe', views.obe, name='obe'),
    path('obe-update', views.obe_update, name='obe_update'),
    path('cam', views.cam, name='cam'),
    path('cam-update', views.cam_update, name='cam_update'),
    path('datasheet', views.datasheet, name='datasheet'),
    path('datasheet-update', views.datasheet_update, name='datasheet_update'),
    path('course-assessment', views.course_assessment, name='course-assessment'),
    path('course-assessment-update', views.course_assessment_update, name='course-assessment_update'),
    path('tos', views.tos, name='tos'),
    path('tos-update', views.tos_update, name='tos_update'),
    path('plo', views.plo, name='plo'),
    path('plo-update', views.plo_update, name='plo_update'),

    path('class-record', views.class_record, name='class_record'),

    path('download-obe', views.download_obe, name='download_obe'),
    path('download-matrix', views.download_matrix, name='download_matrix'),
    path('download-datasheet', views.download_datasheet, name='download_datasheet'),
    path('download-summary', views.download_summary, name='download_summary'),
    path('download-tos', views.download_tos, name='download_tos'),
    path('download-record', views.download_record, name='download_record'),
    path('download-plo', views.download_plo, name='download_plo'),
]
