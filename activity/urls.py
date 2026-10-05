from django.urls import path

from . import views

app_name = 'activity'

urlpatterns = [
    path('', views.drivers_activity, name='drivers_activity'),
    # <path:> so a driver name containing "/" still produces a valid link.
    path('drivers/<path:name>/', views.driver_travel_record, name='driver_travel_record'),
]
