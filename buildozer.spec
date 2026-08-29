# administrador_maquina.spec - Configuración para Buildozer
# Aplicación: Registro de Horas Trabajadas

[app]

# (str) Título de tu aplicación
title = Registro de Horas

# (str) Nombre del paquete (debe ser único)
package.name = registrohoras

# (str) Dominio del paquete (necesario para empaquetar Android/iOS)
package.domain = com.tuempresa

# (str) Archivo principal de la aplicación
# Cambia main.py por tu archivo principal
source.dir = .

# (list) Extensiones de archivos a incluir
source.include_exts = py,png,jpg,kv,atlas,txt,csv

# (list) Incluir archivos específicos
source.include_patterns = administrador_maquina.py,horasapp.kv

# (list) Archivos a excluir
#source.exclude_exts = spec

# (list) Directorios a excluir
source.exclude_dirs = tests, bin, venv, __pycache__, .git

# (list) Patrones de exclusión
#source.exclude_patterns = license,images/*/*.jpg

# (str) Versión de la aplicación
version = 1.0.0

# (list) Requisitos de la aplicación
# IMPORTANTE: Añadimos pandas, openpyxl y numpy (necesario para pandas)
requirements = python3,kivy==2.2.0,pandas,openpyxl,numpy

# (str) Pantalla de presentación (presplash)
#presplash.filename = %(source.dir)s/data/presplash.png

# (str) Icono de la aplicación
#icon.filename = %(source.dir)s/data/icon.png

# (list) Orientaciones soportadas
orientation = portrait

# (list) Servicios a declarar
#services = NAME:ENTRYPOINT_TO_PY

#
# OSX Specific
#

# Author information
# author = Tu Nombre

# Versión de Kivy para OSX
osx.kivy_version = 2.2.0

#
# Android specific
#

# (bool) Aplicación en pantalla completa
fullscreen = 0

# (string) Color de fondo del presplash
#android.presplash_color = #1E88E5

# (str) Icono adaptable (Android API 26+)
#icon.adaptive_foreground.filename = %(source.dir)s/data/icon_fg.png
#icon.adaptive_background.filename = %(source.dir)s/data/icon_bg.png

# (list) Permisos de Android
# Permisos para leer/escribir archivos
android.permissions = android.permission.WRITE_EXTERNAL_STORAGE, android.permission.READ_EXTERNAL_STORAGE, android.permission.INTERNET

# (list) Características (features)
#android.features = 

# (int) API objetivo de Android (recomendado 33 para Play Store)
android.api = 33

# (int) API mínima soportada
android.minapi = 24

# (int) Versión del SDK de Android
#android.sdk = 20

# (str) Versión del NDK de Android
#android.ndk = 23b

# (int) API del NDK
#android.ndk_api = 21

# (str) Directorio del NDK (vacío = descarga automática)
#android.ndk_path =

# (str) Directorio del SDK (vacío = descarga automática)
#android.sdk_path =

# (str) Directorio de ANT
#android.ant_path =

# (bool) Saltar actualización del SDK
# android.skip_update = False

# (bool) Aceptar automáticamente licencias del SDK
# android.accept_sdk_license = False

# (str) Punto de entrada de Android
#android.entrypoint = org.kivy.android.PythonActivity

# (str) Clase Java de la actividad
#android.activity_class_name = org.kivy.android.PythonActivity

# (str) Tema de la app
# android.apptheme = "@android:style/Theme.NoTitleBar"

# (list) Patrón de whitelist
#android.whitelist =

# (bool) Aplicación como launcher
# android.home_app = False

# (list) Archivos .jar a añadir
#android.add_jars =

# (list) Archivos Java a añadir
#android.add_src =

# (list) AARs de Android a añadir
#android.add_aars =

# (list) Archivos para el directorio assets
#android.add_assets =

# (list) Recursos para el directorio res
#android.add_resources =

# (list) Dependencias de Gradle
#android.gradle_dependencies =

# (bool) Habilitar AndroidX
android.enable_androidx = True

# (list) Opciones de compilación Java
# android.add_compile_options = "sourceCompatibility = 1.8", "targetCompatibility = 1.8"

# (list) Repositorios de Gradle
#android.add_gradle_repositories =

# (list) Opciones de empaquetado
#android.add_packaging_options =

# (list) Clases Java como actividades
#android.add_activities =

# (str) Categoría de la consola OUYA
#android.ouya.category = GAME

# (str) Icono de la consola OUYA
#android.ouya.icon.filename = %(source.dir)s/data/ouya_icon.png

# (str) XML para intent filters
#android.manifest.intent_filters =

# (list) Archivos XML para res/xml/
#android.res_xml =

# (str) launchMode de la actividad principal
#android.manifest.launch_mode = standard

# (str) Orientación de la pantalla
#android.manifest.orientation = fullSensor

# (list) Librerías adicionales para Android
#android.add_libs_armeabi =
#android.add_libs_armeabi_v7a =
#android.add_libs_arm64_v8a =
#android.add_libs_x86 =
#android.add_libs_mips =

# (bool) Mantener pantalla encendida
#android.wakelock = False

# (list) Meta-data de la aplicación
#android.meta_data =

# (list) Referencias de librerías de Android
#android.library_references =

# (list) Librerías compartidas de Android
#android.uses_library =

# (str) Filtros de logcat
#android.logcat_filters = *:S python:D

# (bool) Mostrar solo log del PID de la actividad
#android.logcat_pid_only = False

# (str) Argumentos adicionales de adb
#android.adb_args =

# (bool) Copiar librerías en lugar de crear libpymodules.so
#android.copy_libs = 1

# (list) Arquitecturas de Android a compilar
android.archs = arm64-v8a, armeabi-v7a

# (int) Versión numérica (override)
# android.numeric_version = 1

# (bool) Habilitar backup automático
android.allow_backup = True

# (str) Reglas de backup personalizadas
# android.backup_rules =

# (str) Placeholders para AndroidManifest.xml
# android.manifest_placeholders = [:]

# (bool) Saltar compilación byte de archivos .py
# android.no-byte-compile-python = False

# (str) Formato de empaquetado release (aab o apk)
# android.release_artifact = apk

# (str) Formato de empaquetado debug (apk o aar)
# android.debug_artifact = apk

# (str) Recorte de pantalla (display cutout)
#android.display_cutout = never

#
# Python for Android (p4a) specific
#

# (str) URL de python-for-android
#p4a.url =

# (str) Fork de python-for-android
#p4a.fork = kivy

# (str) Rama de python-for-android
#p4a.branch = master

# (str) Commit específico de python-for-android
#p4a.commit = HEAD

# (str) Directorio de clonación de p4a
#p4a.source_dir =

# (str) Directorio de recetas locales
#p4a.local_recipes =

# (str) Hook para p4a
#p4a.hook =

# (str) Bootstrap para Android
# p4a.bootstrap = sdl2

# (int) Puerto para p4a
#p4a.port =

# (bool) Usar setup.py
#p4a.setup_py = false

# (str) Argumentos extra para p4a
#p4a.extra_args =

#
# iOS specific
#

# (str) Directorio de kivy-ios
#ios.kivy_ios_dir = ../kivy-ios

# (str) URL de kivy-ios
ios.kivy_ios_url = https://github.com/kivy/kivy-ios
ios.kivy_ios_branch = master

# (str) Directorio de ios-deploy
#ios.ios_deploy_dir = ../ios_deploy

# (str) URL de ios-deploy
ios.ios_deploy_url = https://github.com/phonegap/ios-deploy
ios.ios_deploy_branch = 1.12.2

# (bool) Firmar código
ios.codesign.allowed = false

# (str) Certificado para debug
#ios.codesign.debug = "iPhone Developer: <lastname> <firstname> (<hexstring>)"

# (str) Team de desarrollo para debug
#ios.codesign.development_team.debug = <hexstring>

# (str) Certificado para release
#ios.codesign.release = %(ios.codesign.debug)s

# (str) Team de desarrollo para release
#ios.codesign.development_team.release = <hexstring>

# (str) Justificación para acceso a medios
#ios.media_usage_description =

# (str) Justificación para red local
#ios.local_network_usage_description =

# (str) Justificación para cámara
#ios.camera_usage_description =

# (bool) Control de StatusBar
# ios.viewcontroller_based_statusbar_appearance = False

# (str) Extensiones de la app
#ios.app_extensions =

# (str) URL del .ipa
#ios.manifest.app_url =

# (str) URL del icono (57x57)
#ios.manifest.display_image_url =

# (str) URL del icono grande (512x512)
#ios.manifest.full_size_image_url =

[buildozer]

# (int) Nivel de log (0=errores, 1=info, 2=debug)
log_level = 2

# (int) Mostrar advertencia si se ejecuta como root
warn_on_root = 1

# (str) Directorio de build
# build_dir = ./.buildozer

# (str) Directorio de salida (.apk, .aab, .ipa)
# bin_dir = ./bin