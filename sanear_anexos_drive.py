"""
Saneamiento de anexos — Sistema de Incidencias DFC
====================================================
Se ejecuta UNA SOLA VEZ desde tu máquina local.

Hace dos cosas:
  1. Te agrega como miembro de la unidad compartida (para que los links
     del Sheet te sigan abriendo después de quitar el acceso público).
  2. Revoca el permiso "anyone with link" de TODOS los archivos ya
     subidos a la carpeta de anexos.

Requisitos:
    pip install google-api-python-client google-auth toml

Uso:
    1. Edita las 2 constantes de abajo.
    2. Corre el script DESDE LA RAÍZ del proyecto (donde está .streamlit/):
       python sanear_anexos_drive.py
"""

import toml
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

# ── CONFIGURA ESTAS 2 LÍNEAS ─────────────────────────────────────
CARPETA_ANEXOS_ID = "1LnQjrhjEKgKxFTJD8USLoiCpCKQHfOQC"             # mismo valor que DRIVE_ANEXOS_FOLDER en la app
TU_CORREO         = "martin.carrizalez@jalisco.gob.mx"    # correo que debe poder ver los anexos
# ─────────────────────────────────────────────────────────────────

secrets = toml.load(".streamlit/secrets.toml")
creds = Credentials.from_service_account_info(
    secrets["gcp_service_account"], scopes=["https://www.googleapis.com/auth/drive"]
)
drive = build("drive", "v3", credentials=creds)

# ── PASO 1: agregarte como miembro de la unidad compartida ──────
print("Paso 1 — Membresía de la unidad compartida")
carpeta = drive.files().get(
    fileId=CARPETA_ANEXOS_ID, fields="id,name,driveId", supportsAllDrives=True
).execute()
drive_id = carpeta.get("driveId")

if drive_id:
    try:
        drive.permissions().create(
            fileId=drive_id,
            body={"type": "user", "role": "fileOrganizer", "emailAddress": TU_CORREO},
            supportsAllDrives=True,
            sendNotificationEmail=False,
        ).execute()
        print(f"  ✅ {TU_CORREO} agregado a la unidad compartida como Gestor de contenido.")
    except Exception as e:
        if "already" in str(e).lower() or "duplicate" in str(e).lower():
            print(f"  ℹ️ {TU_CORREO} ya era miembro. OK.")
        else:
            print(f"  ⚠️ No se pudo agregar a la unidad: {e}")
            print("     (Si falla, comparte la carpeta manualmente o pide al admin de Workspace.)")
else:
    # La carpeta está en "Mi unidad" de la service account, no en unidad compartida
    print("  ℹ️ La carpeta NO está en una unidad compartida — comparto la carpeta directamente.")
    drive.permissions().create(
        fileId=CARPETA_ANEXOS_ID,
        body={"type": "user", "role": "reader", "emailAddress": TU_CORREO},
        sendNotificationEmail=False,
    ).execute()
    print(f"  ✅ Carpeta compartida con {TU_CORREO} (los archivos heredan el permiso).")

# ── PASO 2: revocar 'anyone' de todos los archivos existentes ───
print("\nPaso 2 — Revocando acceso público de archivos existentes")
revocados, sin_publico, errores = 0, 0, 0
page_token = None

while True:
    resp = drive.files().list(
        q=f"'{CARPETA_ANEXOS_ID}' in parents and trashed = false",
        fields="nextPageToken, files(id, name)",
        pageSize=100,
        pageToken=page_token,
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
    ).execute()

    for f in resp.get("files", []):
        try:
            perms = drive.permissions().list(
                fileId=f["id"], fields="permissions(id,type)", supportsAllDrives=True
            ).execute().get("permissions", [])
            perm_anyone = next((p for p in perms if p["type"] == "anyone"), None)
            if perm_anyone:
                drive.permissions().delete(
                    fileId=f["id"], permissionId=perm_anyone["id"], supportsAllDrives=True
                ).execute()
                print(f"  🔒 {f['name']} — acceso público revocado")
                revocados += 1
            else:
                sin_publico += 1
        except Exception as e:
            print(f"  ⚠️ {f['name']}: {e}")
            errores += 1

    page_token = resp.get("nextPageToken")
    if not page_token:
        break

print(f"\nResumen: {revocados} revocados · {sin_publico} ya estaban privados · {errores} errores")
print("\nVerificación final: abre un link de LINK_ANEXO en ventana de incógnito.")
print("Debe decir 'Solicitar acceso'. Con tu sesión institucional debe abrir normal.")
