from flask import Flask, render_template, request, redirect, url_for, session, flash
import firebase_admin
from firebase_admin import credentials, firestore, storage 
from PIL import Image 
import io
from datetime import datetime
import re
import uuid 
import os
import json
import base64
import threading # IMPORTANTE: Para el envío en segundo plano
from gmail_api import send_email

# 1. Configuración de Firebase (local: credenciales.json | Render: FIREBASE_CREDENTIALS_B64)
def _decode_service_account_b64(b64: str) -> dict:
    """Acepta base64 pegado desde el portapapeles (espacios/saltos de línea)."""
    raw = "".join(b64.split())
    pad = (-len(raw)) % 4
    if pad:
        raw += "=" * pad
    decoded = base64.b64decode(raw)
    return json.loads(decoded.decode("utf-8"))


def _firebase_certificate():
    cred_path = "credenciales.json"
    if os.path.isfile(cred_path):
        return credentials.Certificate(cred_path)
    b64 = os.environ.get("FIREBASE_CREDENTIALS_B64")
    if b64:
        try:
            data = _decode_service_account_b64(b64)
            return credentials.Certificate(data)
        except (json.JSONDecodeError, ValueError, UnicodeDecodeError) as e:
            raise RuntimeError(
                "Firebase: FIREBASE_CREDENTIALS_B64 no es un JSON válido en base64. "
                "Vuelve a generarlo en PowerShell y pégalo en Render en una sola línea, sin comillas."
            ) from e
    raise RuntimeError(
        "Firebase: crea credenciales.json en la raíz o define la variable de entorno FIREBASE_CREDENTIALS_B64."
    )

if not firebase_admin._apps:
    cred = _firebase_certificate()
    firebase_admin.initialize_app(cred, {
        'storageBucket': 'gestion-charat-admin.firebasestorage.app' 
    })

db = firestore.client()
bucket = storage.bucket() 

app = Flask(__name__)
app.secret_key = "muni_charat_2026_secure_key"


@app.get("/health")
def health():
    """Para comprobar que el proceso responde (Render / balanceadores)."""
    return "ok", 200


def optimizar_imagen(archivo_foto):
    img = Image.open(archivo_foto)
    if img.mode in ("RGBA", "P"):
        img = img.convert("RGB")
    img.thumbnail((800, 800))
    buffer = io.BytesIO()
    img.save(buffer, format="JPEG", quality=60, optimize=True)
    buffer.seek(0)
    return buffer

# FUNCIÓN QUE CORRE EN SEGUNDO PLANO
def enviar_correo_async(app_context, categoria, nombre, caserio, descripcion):
    with app_context:
        try:
            config_doc = db.collection('configuracion').document('email_settings').get()
            if config_doc.exists:
                conf = config_doc.to_dict()
                correo_receptor = conf.get('correo_receptor')
                if correo_receptor:
                    asunto = f"MuniCharat: {categoria}"
                    cuerpo = f"Nuevo reporte de {nombre}.\nLugar: {caserio}\nDetalle: {descripcion}"
                    send_email(to_email=correo_receptor, subject=asunto, body_text=cuerpo)
                    print("LOG: Correo enviado exitosamente (Gmail API).")
        except Exception as e:
            print(f"LOG ERROR CORREO: {e}")

@app.route('/')
def index():
    try:
        sectores_docs = db.collection('sectores').order_by('nombre').stream()
        lista_sectores = [{"nombre": doc.to_dict().get('nombre', 'N/A'), "tipo": doc.to_dict().get('tipo', 'Sector')} for doc in sectores_docs]
    except Exception:
        lista_sectores = [] 
    return render_template('index.html', sectores=lista_sectores)

@app.route('/enviar', methods=['POST'])
def enviar():
    try:
        dni = request.form.get('dni')
        nombre = request.form.get('nombre')
        categoria = request.form.get('categoria')
        descripcion = request.form.get('descripcion')
        caserio = request.form.get('caserio')
        
        if not re.match(r"^\d{8}$", dni): return "Error DNI", 400

        lista_urls_fotos = []
        archivos = request.files.getlist('foto') 
        for archivo in archivos:
            if archivo and archivo.filename != '':
                nombre_unico = f"incidencias/{uuid.uuid4()}.jpg"
                blob = bucket.blob(nombre_unico)
                foto_comprimida = optimizar_imagen(archivo)
                blob.upload_from_file(foto_comprimida, content_type='image/jpeg')
                blob.make_public()
                lista_urls_fotos.append(blob.public_url)

        nueva_incidencia = {
            "informante": {"nombre": nombre, "dni": dni},
            "ubicacion": {"caserio": caserio, "lat": request.form.get('latitud'), "lng": request.form.get('longitud')},
            "detalle": {"categoria": categoria, "descripcion": descripcion, "fotos": lista_urls_fotos},
            "gestion": {"estado": "Pendiente", "fecha_registro": datetime.now(), "ultima_actualizacion": datetime.now()}
        }
        
        # GUARDADO EN FIREBASE
        db.collection('incidencias').add(nueva_incidencia)

        # DISPARAR CORREO EN HILO SEPARADO
        threading.Thread(target=enviar_correo_async, 
                         args=(app.app_context(), categoria, nombre, caserio, descripcion)).start()

        return render_template('exito.html')

    except Exception as e:
        print(f"ERROR CRÍTICO: {e}")
        return f"Error: {e}", 500

# --- RUTAS ADMINISTRATIVAS ---

@app.route('/admin')
def admin():
    if not session.get('logged_in'): return redirect(url_for('login'))
    reportes = [doc.to_dict() | {"id": doc.id} for doc in db.collection('incidencias').stream()]
    reportes.sort(key=lambda x: x['gestion'].get('fecha_registro', datetime.now()), reverse=True)
    sectores = [doc.to_dict() | {"id": doc.id} for doc in db.collection('sectores').order_by('nombre').stream()]
    config_actual = {}
    usuarios = []
    if session.get('rol') == 'admin':
        usuarios = [doc.to_dict() | {"id": doc.id} for doc in db.collection('usuarios').stream()]
        conf_doc = db.collection('configuracion').document('email_settings').get()
        if conf_doc.exists: config_actual = conf_doc.to_dict()
    return render_template('admin.html', reportes=reportes, sectores=sectores, usuarios=usuarios, config=config_actual)

@app.route('/guardar_config_correo', methods=['POST'])
def guardar_config_correo():
    if not session.get('logged_in') or session.get('rol') != 'admin': return "No autorizado", 403
    db.collection('configuracion').document('email_settings').set({
        "correo_receptor": request.form.get('correo_receptor').strip(),
        "ultima_actualizacion": datetime.now()
    })
    return redirect(url_for('admin'))

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        u, p = request.form.get('usuario'), request.form.get('password')
        q = db.collection('usuarios').where('usuario', '==', u).where('password', '==', p).limit(1).get()
        if q:
            session.update({'logged_in': True, 'user': q[0].to_dict()['usuario'], 'rol': q[0].to_dict().get('rol', 'operador')})
            return redirect(url_for('admin'))
        return render_template('login.html', error="Credenciales incorrectas")
    return render_template('login.html')

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)