from django.urls import path

from . import views

app_name = 'activity'

urlpatterns = [
    path('', views.drivers_activity, name='drivers_activity'),
    path('import/', views.import_travel_report, name='import_travel_report'),
    path('export/drivers/', views.export_all_travel_records, name='export_all_travel_records'),
    path('export/drivers/<path:name>/', views.export_travel_record, name='export_travel_record'),
    # <path:> so a driver name containing "/" still produces a valid link.
    path('drivers/<path:name>/', views.driver_travel_record, name='driver_travel_record'),
]
