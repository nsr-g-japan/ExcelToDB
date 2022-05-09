from django.urls import path
from .views import  *
urlpatterns = [
    path('homepage', homepage),
    path('sheetname', sheetname),

]
