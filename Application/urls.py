from django.urls import path
from Application import views

app_name = "fileapp"

urlpatterns = [
    path('', views.index),
    path("roll-marksheet", views.roll_marksheet, name = "roll-marksheet"),
    path("concise-marksheet", views.concise_marksheet, name = "concise-marksheet"),
    path("email", views.send_email, name = "email"),
]