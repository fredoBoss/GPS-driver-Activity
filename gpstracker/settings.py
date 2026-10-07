"""
Django settings for gpstracker project.

The site has no database models: every page reads the Excel activity log
(ACTIVITY_LOG_PATH) directly, so there is nothing to migrate.
"""

import os
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent

# Development defaults. Set DJANGO_SECRET_KEY, DJANGO_DEBUG=0 and DJANGO_ALLOWED_HOSTS
# before serving this to anyone else.
SECRET_KEY = os.environ.get(
    'DJANGO_SECRET_KEY',
    'django-insecure-8h&h%49j0#5ne0xzzr$3c%d%b3te4--c5pmekygtiz*t_1jh76',
)
DEBUG = os.environ.get('DJANGO_DEBUG', '1') != '0'
ALLOWED_HOSTS = [h for h in os.environ.get('DJANGO_ALLOWED_HOSTS', '').split(',') if h]

# The Excel workbook the pages read (Driver_Activity_Log.xlsx in the project root by default).
ACTIVITY_LOG_PATH = Path(os.environ.get('ACTIVITY_LOG_PATH', BASE_DIR / 'Driver_Activity_Log.xlsx'))


INSTALLED_APPS = [
    'django.contrib.messages',
    'django.contrib.staticfiles',
    'activity',
]

MIDDLEWARE = [
    'django.middleware.security.SecurityMiddleware',
    'django.middleware.common.CommonMiddleware',
    'django.middleware.csrf.CsrfViewMiddleware',
    'django.contrib.messages.middleware.MessageMiddleware',
    'django.middleware.clickjacking.XFrameOptionsMiddleware',
]

# Cookie storage keeps "imported N trips" messages working without sessions or a database.
MESSAGE_STORAGE = 'django.contrib.messages.storage.cookie.CookieStorage'

ROOT_URLCONF = 'gpstracker.urls'

TEMPLATES = [
    {
        'BACKEND': 'django.template.backends.django.DjangoTemplates',
        'DIRS': [],
        'APP_DIRS': True,
        'OPTIONS': {
            'context_processors': [
                'django.template.context_processors.request',
                'django.contrib.messages.context_processors.messages',
            ],
        },
    },
]

WSGI_APPLICATION = 'gpstracker.wsgi.application'

DATABASES = {
    'default': {
        'ENGINE': 'django.db.backends.sqlite3',
        'NAME': BASE_DIR / 'db.sqlite3',
    }
}


LANGUAGE_CODE = 'en-us'

TIME_ZONE = 'UTC'

USE_I18N = True

USE_TZ = True


STATIC_URL = 'static/'

DEFAULT_AUTO_FIELD = 'django.db.models.BigAutoField'
