"""URL configuration for gpstracker project."""
from django.urls import include, path

urlpatterns = [
    path('', include('activity.urls')),
]
