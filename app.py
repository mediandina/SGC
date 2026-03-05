from flask import Flask, render_template, request, redirect, url_for, jsonify, session, send_file
from werkzeug.security import generate_password_hash, check_password_hash
from openpyxl import load_workbook
from openpyxl.styles import Border, Side, Font, Alignment
from datetime import datetime, date
from email.message import EmailMessage


import pandas as pd
import os
import re
import smtplib


app = Flask(__name__)

app.secret_key = "clave-secreta-sgf"  # Cambia esto por una clave segura en producción

EXCEL_FILE = "cupos.xlsx"

# -------------------------------------------------------
# CONFIGURACIÓN DE CORREO — edita estas dos líneas
# -------------------------------------------------------
EMAIL_REMITENTE = "multimediandina@gmail.com"        # ← tu correo Gmail
EMAIL_APP_PASSWORD = "tadg fwgb wett wxlj"    # ← App Password de Google (16 caracteres)
# -------------------------------------------------------


# ---------- CREAR EXCEL SI NO EXISTE ----------
COLUMNAS = [
    "Fecha",
    "Nombre del conductor",
    "Tipo de vehículo",
    "Cupo",
    "Proveedor",
    "Telefono",
    "Correo",
    "Placa",
    "Kilos aproximados",
    "Pacas"
]

USUARIOS_FILE = "usuarios.xlsx"

# ---------- CREAR TABLA DE USUARIOS ----------
COLUMNAS_USUARIOS = [
    "Nombre",
    "Telefono",
    "Proveedor",
    "Password"
]

if not os.path.exists(USUARIOS_FILE):
    pd.DataFrame(columns=COLUMNAS_USUARIOS).to_excel(USUARIOS_FILE, index=False)

if not os.path.exists(EXCEL_FILE):
    pd.DataFrame(columns=COLUMNAS).to_excel(EXCEL_FILE, index=False)


def es_excel_cupos(path):
    try:
        df = pd.read_excel(path)
        return list(df.columns) == COLUMNAS
    except:
        return False


def asegurar_excel():
    candidatos = []
    for archivo in os.listdir("."):
        if archivo.endswith(".xlsx"):
            if es_excel_cupos(archivo):
                try:
                    df = pd.read_excel(archivo)
                    filas = len(df)
                    candidatos.append((archivo, filas))
                except:
                    pass

    if candidatos:
        candidatos.sort(key=lambda x: x[1], reverse=True)
        archivo_principal = candidatos[0][0]
        if archivo_principal != EXCEL_FILE:
            if os.path.exists(EXCEL_FILE):
                os.remove(EXCEL_FILE)
            os.rename(archivo_principal, EXCEL_FILE)
        return

    pd.DataFrame(columns=COLUMNAS).to_excel(EXCEL_FILE, index=False)


def formatear_excel():
    wb = load_workbook(EXCEL_FILE)
    ws = wb.active

    borde_fino = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )

    for col in range(1, ws.max_column + 1):
        celda = ws.cell(row=1, column=col)
        celda.font = Font(bold=True)
        celda.alignment = Alignment(horizontal="center", vertical="center")
        celda.border = borde_fino
        ws.column_dimensions[celda.column_letter].width = 22

    for row in range(2, ws.max_row + 1):
        for col in range(1, ws.max_column + 1):
            celda = ws.cell(row=row, column=col)
            celda.border = borde_fino
            celda.alignment = Alignment(horizontal="center", vertical="center")
            if col == 6 or celda.column_letter == "F":
                celda.number_format = "@"

    wb.save(EXCEL_FILE)


def formatear_excel_usuarios():
    wb = load_workbook(USUARIOS_FILE)
    ws = wb.active

    borde_fino = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )

    for col in range(1, ws.max_column + 1):
        celda = ws.cell(row=1, column=col)
        celda.font = Font(bold=True)
        celda.alignment = Alignment(horizontal="center", vertical="center")
        celda.border = borde_fino
        ws.column_dimensions[celda.column_letter].width = 22

    for row in range(2, ws.max_row + 1):
        for col in range(1, ws.max_column + 1):
            celda = ws.cell(row=row, column=col)
            celda.border = borde_fino
            celda.alignment = Alignment(horizontal="center", vertical="center")
            if col == 2 or celda.column_letter == "B":
                celda.number_format = "@"

    wb.save(USUARIOS_FILE)


# =======================================================
# FUNCIÓN DE ENVÍO DE CORREO DE CONFIRMACIÓN
# =======================================================
def enviar_correo_confirmacion(destinatario, nombre, fecha, cupo):
    """
    Envía un correo HTML de confirmación de cupo al conductor.
    Retorna True si se envió correctamente, False si hubo error.
    El error NO interrumpe el guardado del cupo.
    """
    try:
        meses = {
            1: "enero", 2: "febrero", 3: "marzo", 4: "abril",
            5: "mayo", 6: "junio", 7: "julio", 8: "agosto",
            9: "septiembre", 10: "octubre", 11: "noviembre", 12: "diciembre"
        }
        dias_semana = {
            0: "lunes", 1: "martes", 2: "miércoles", 3: "jueves",
            4: "viernes", 5: "sábado", 6: "domingo"
        }
        fecha_obj = datetime.strptime(fecha, "%Y-%m-%d")
        fecha_legible = (
            f"{dias_semana[fecha_obj.weekday()]} "
            f"{fecha_obj.day} de "
            f"{meses[fecha_obj.month]} de "
            f"{fecha_obj.year}"
        )

        asunto = f"Confirmacion de cupo #{cupo} - {fecha_legible}"

        cuerpo_html = f"""
<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
</head>
<body style="margin:0; padding:0; background-color:#f5f5f3; font-family:'Segoe UI', Arial, sans-serif;">

  <table width="100%" cellpadding="0" cellspacing="0" style="background-color:#f5f5f3; padding: 32px 16px;">
    <tr>
      <td align="center">
        <table width="600" cellpadding="0" cellspacing="0" style="max-width:600px; width:100%; background-color:#ffffff; border-radius:12px; overflow:hidden; box-shadow: 0 4px 20px rgba(0,0,0,0.08);">

          <!-- ENCABEZADO VERDE -->
          <tr>
            <td style="background: linear-gradient(135deg, #6bbf2d 0%, #5aa524 100%); padding: 32px 40px; text-align:center;">
              <h1 style="margin:0; color:#ffffff; font-size:26px; font-weight:700; letter-spacing:-0.5px;">
                Descargue Andina
              </h1>
              <p style="margin:8px 0 0; color:rgba(255,255,255,0.88); font-size:14px;">
                Sistema de Gestion de Cupos
              </p>
            </td>
          </tr>

          <!-- BANNER CONFIRMACION -->
          <tr>
            <td style="background-color:#f0f9e8; padding: 20px 40px; text-align:center; border-bottom: 3px solid #6bbf2d;">
              <p style="margin:0 0 6px; font-size:36px;">&#10003;</p>
              <h2 style="margin:0 0 4px; color:#1f2937; font-size:20px; font-weight:700;">
                Cupo confirmado exitosamente
              </h2>
              <p style="margin:0; color:#4b5563; font-size:14px;">
                Su agendamiento ha sido registrado en el sistema.
              </p>
            </td>
          </tr>

          <!-- DETALLE -->
          <tr>
            <td style="padding: 32px 40px;">
              <p style="margin:0 0 20px; color:#1f2937; font-size:15px;">
                Hola, <strong>{nombre}</strong>. A continuacion encontrara el detalle de su cupo agendado:
              </p>

              <table width="100%" cellpadding="0" cellspacing="0" style="background-color:#f9fafb; border-radius:8px; border:1px solid #e5e7eb;">
                <tr>
                  <td style="padding:14px 20px; border-bottom:1px solid #e5e7eb;">
                    <span style="color:#6b7280; font-size:11px; font-weight:700; text-transform:uppercase; letter-spacing:0.6px; display:block; margin-bottom:4px;">Fecha asignada</span>
                    <span style="color:#1f2937; font-size:16px; font-weight:700; text-transform:capitalize;">{fecha_legible}</span>
                  </td>
                </tr>
                <tr>
                  <td style="padding:14px 20px; border-bottom:1px solid #e5e7eb;">
                    <span style="color:#6b7280; font-size:11px; font-weight:700; text-transform:uppercase; letter-spacing:0.6px; display:block; margin-bottom:6px;">Numero de cupo</span>
                    <span style="display:inline-block; background-color:#6bbf2d; color:#ffffff; font-size:22px; font-weight:800; padding:6px 24px; border-radius:6px;">
                      Cupo {cupo}
                    </span>
                  </td>
                </tr>
                <tr>
                  <td style="padding:14px 20px;">
                    <span style="color:#6b7280; font-size:11px; font-weight:700; text-transform:uppercase; letter-spacing:0.6px; display:block; margin-bottom:4px;">Conductor</span>
                    <span style="color:#1f2937; font-size:15px; font-weight:600;">{nombre}</span>
                  </td>
                </tr>
              </table>

              <!-- AVISO LLAMADA -->
              <table width="100%" cellpadding="0" cellspacing="0" style="margin-top:20px; border-left:4px solid #c7b404; background-color:#fffbf0; border-radius:4px;">
                <tr>
                  <td style="padding:14px 16px;">
                    <p style="margin:0; color:#1f2937; font-size:14px; line-height:1.6;">
                      <strong>Recuerde:</strong> el dia del cupo asignado, el personal de Descargue Andina
                      se comunicara <strong>telefonicamente</strong> con usted para indicarle que puede
                      acercarse al area de descarga.
                    </p>
                  </td>
                </tr>
              </table>
            </td>
          </tr>

          <!-- REGLAS -->
          <tr>
            <td style="padding: 0 40px 32px;">
              <h3 style="margin:0 0 14px; color:#1f2937; font-size:15px; font-weight:700; border-bottom:2px solid #e5e7eb; padding-bottom:8px;">
                Reglas importantes
              </h3>
              <table width="100%" cellpadding="0" cellspacing="0">
                <tr><td style="padding:8px 0; border-bottom:1px solid #f3f4f6; vertical-align:top;">
                  <span style="color:#6bbf2d; font-weight:700; margin-right:8px;">&#10003;</span>
                  <span style="color:#374151; font-size:13.5px; line-height:1.6;">Todos los datos solicitados deben ser diligenciados correctamente.</span>
                </td></tr>
                <tr><td style="padding:8px 0; border-bottom:1px solid #f3f4f6; vertical-align:top;">
                  <span style="color:#6bbf2d; font-weight:700; margin-right:8px;">&#10003;</span>
                  <span style="color:#374151; font-size:13.5px; line-height:1.6;">En caso de no poder asistir el dia programado, comuniquese con el area de despacho con la debida anticipacion.</span>
                </td></tr>
                <tr><td style="padding:8px 0; border-bottom:1px solid #f3f4f6; vertical-align:top;">
                  <span style="color:#6bbf2d; font-weight:700; margin-right:8px;">&#10003;</span>
                  <span style="color:#374151; font-size:13.5px; line-height:1.6;">Para el ingreso es obligatorio contar con el ticket de bascula (autorizada) correspondiente al peso inicial debidamente autorizado.</span>
                </td></tr>
                <tr><td style="padding:8px 0; vertical-align:top;">
                  <span style="color:#6bbf2d; font-weight:700; margin-right:8px;">&#10003;</span>
                  <span style="color:#374151; font-size:13.5px; line-height:1.6;">Por favor comprobar que los datos registrados sean iguales a los datos solicitados en el formulario.</span>
                </td></tr>
              </table>
            </td>
          </tr>

          <!-- FOOTER -->
          <tr>
            <td style="background-color:#111213; padding:24px 40px; text-align:center; border-radius:0 0 12px 12px;">
              <p style="margin:0 0 6px; color:#ffffff; font-size:14px; font-weight:600;">
                Descargue Andina &mdash; Servicio al cliente
              </p>
              <p style="margin:0; color:#9ca3af; font-size:12px;">
                +57 313 8893206 &nbsp;|&nbsp;
                <a href="mailto:mprimas@corrugadosandina.com.co" style="color:#6bbf2d; text-decoration:none;">
                  mprimas@corrugadosandina.com.co
                </a>
              </p>
              <p style="margin:12px 0 0; color:#6b7280; font-size:11px;">
                Este es un correo automatico, por favor no responda a este mensaje.
              </p>
            </td>
          </tr>

        </table>
      </td>
    </tr>
  </table>

</body>
</html>
"""

        msg = EmailMessage()
        msg["Subject"] = asunto
        msg["From"] = f"Descargue Andina <{EMAIL_REMITENTE}>"
        msg["To"] = destinatario
        # Texto plano como fallback para clientes sin HTML
        msg.set_content(
            f"Cupo confirmado: #{cupo} para el {fecha_legible}.\n"
            f"Conductor: {nombre}\n\n"
            "Este correo contiene formato HTML. Abralo en un cliente compatible para verlo correctamente."
        )
        msg.add_alternative(cuerpo_html, subtype="html")

        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
            smtp.login(EMAIL_REMITENTE, EMAIL_APP_PASSWORD)
            smtp.send_message(msg)

        print(f"[EMAIL OK] Confirmacion enviada a {destinatario}")
        return True

    except Exception as e:
        print(f"[EMAIL ERROR] No se pudo enviar a {destinatario}: {e}")
        return False


# ---------- Registro ----------
@app.route("/registro", methods=["GET", "POST"])
def registro():
    if "usuario" in session:
        return redirect(url_for("sistema"))

    if request.method == "POST":
        try:
            df = pd.read_excel(USUARIOS_FILE)

            telefono = request.form.get("telefono", "").strip()
            telefono = re.sub(r"\D", "", telefono)

            if "Telefono" in df.columns:
                df["Telefono"] = df["Telefono"].astype(str).str.replace(r"\.0$", "", regex=True).str.strip()
            else:
                df["Telefono"] = df.get("Telefono", pd.Series(dtype=str))

            if telefono in df["Telefono"].values:
                return render_template("error.html", mensaje="El telefono ya esta registrado")

            nuevo = {
                "Nombre": request.form.get("nombre"),
                "Telefono": telefono,
                "Proveedor": request.form.get("proveedor"),
                "Password": generate_password_hash(request.form.get("password"))
            }

            df = pd.concat([df, pd.DataFrame([nuevo])], ignore_index=True)
            df.to_excel(USUARIOS_FILE, index=False)
            formatear_excel_usuarios()

            session["usuario"] = telefono
            return redirect(url_for("sistema"))

        except PermissionError:
            return jsonify({
                "error": "archivo_bloqueado",
                "mensaje": "El sistema se encuentra actualmente ocupado. Por favor, espere un momento."
            }), 503

        except Exception as e:
            print("ERROR:", e)
            return render_template("error.html", mensaje="Ocurrio un error inesperado. Intente mas tarde."), 500

    return render_template("formulario.html")


# Credenciales temporales de administrador
ADMIN_TEL = "1234567890"
ADMIN_PASS = "admin123"


@app.route("/login", methods=["GET", "POST"])
def login():
    if "usuario" in session:
        return redirect(url_for("sistema"))

    if request.method == "POST":
        telefono = request.form.get("telefono", "").strip()
        telefono = re.sub(r"\D", "", telefono)
        password = request.form.get("password", "")

        if telefono == ADMIN_TEL and password == ADMIN_PASS:
            session["usuario"] = telefono
            session["is_admin"] = True
            return redirect(url_for("admin_panel"))

        df = pd.read_excel(USUARIOS_FILE)
        if "Telefono" in df.columns:
            df["Telefono"] = df["Telefono"].astype(str).str.replace(r"\.0$", "", regex=True).str.strip()

        user = df[df["Telefono"] == telefono]

        if user.empty:
            return render_template("error.html", mensaje="Usuario no encontrado")

        if not check_password_hash(user.iloc[0]["Password"], password):
            return render_template("error.html", mensaje="Contrasena incorrecta")

        session["usuario"] = telefono
        session.pop("is_admin", None)
        return redirect(url_for("sistema"))

    return render_template("login.html")


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


# ---------- VISTA PRINCIPAL ----------
@app.route("/")
def sistema():
    if "usuario" not in session:
        return redirect(url_for("login"))
    if session.get("is_admin"):
        return redirect(url_for("admin_panel"))
    return render_template("sistema.html")


# ---------- PANEL ADMIN ----------
@app.route("/admin")
def admin_panel():
    if "usuario" not in session or not session.get("is_admin"):
        return redirect(url_for("login"))
    try:
        asegurar_excel()
        if not os.path.exists(EXCEL_FILE):
            return render_template("admin.html", columnas=[], filas=[])
        df = pd.read_excel(EXCEL_FILE).fillna("")
        columnas = df.columns.tolist()
        filas = df.to_dict(orient="records")
        return render_template("admin.html", columnas=columnas, filas=filas)
    except Exception as e:
        print("ERROR admin_panel:", e)
        return render_template("error.html", mensaje="No se pudo cargar los datos del sistema"), 500


@app.route("/download_cupos")
def download_cupos():
    if "usuario" not in session or not session.get("is_admin"):
        return redirect(url_for("login"))
    try:
        asegurar_excel()
        if not os.path.exists(EXCEL_FILE):
            return render_template("error.html", mensaje="Archivo de cupos no encontrado"), 404
        return send_file(EXCEL_FILE, as_attachment=True)
    except Exception as e:
        print("ERROR download_cupos:", e)
        return render_template("error.html", mensaje="No se pudo descargar el archivo"), 500


# ---------- CUPOS OCUPADOS ----------
@app.route("/cupos_ocupados")
def cupos_ocupados():
    if "usuario" not in session:
        return jsonify({"ocupadas": []})
    try:
        asegurar_excel()
        fecha = request.args.get("fecha")
        ocupadas = []
        if fecha:
            df = pd.read_excel(EXCEL_FILE)
            df_fecha = df[df["Fecha"] == fecha]
            ocupadas = df_fecha["Cupo"].tolist()
        return jsonify({"ocupadas": ocupadas})
    except Exception:
        return jsonify({"ocupadas": []})


# ---------- GUARDAR CUPO ----------
@app.route("/guardar", methods=["POST"])
def guardar_cupo():
    if "usuario" not in session:
        return redirect(url_for("login"))
    try:
        asegurar_excel()

        errores = validar_datos(request.form)
        if errores:
            return render_template("error.html", mensaje="Datos invalidos enviados al sistema"), 400

        fecha = request.form.get("fecha")
        cupo = int(request.form.get("cupo"))

        fecha_cupo = datetime.strptime(fecha, "%Y-%m-%d").date()
        hoy = date.today()

        if fecha_cupo < hoy:
            return render_template("error.html", mensaje="No se pueden registrar cupos para fechas pasadas"), 400

        if fecha_cupo.weekday() == 6:
            return render_template("error.html", mensaje="No se pueden registrar cupos los domingos"), 400

        max_cupos = 12 if fecha_cupo.weekday() <= 4 else 5

        if cupo < 1 or cupo > max_cupos:
            return render_template("error.html", mensaje="El numero de cupo no es valido para ese dia"), 400

        df = pd.read_excel(EXCEL_FILE)

        if ((df["Fecha"] == fecha) & (df["Cupo"] == cupo)).any():
            return render_template("error.html", mensaje="El cupo seleccionado ya esta ocupado"), 409

        nombre = request.form.get("nombre")
        correo = request.form.get("email")

        data = {
            "Fecha": fecha,
            "Nombre del conductor": nombre,
            "Tipo de vehículo": request.form.get("tipo"),
            "Cupo": cupo,
            "Proveedor": request.form.get("proveedor"),
            "Telefono": str(request.form.get("telefono")),
            "Correo": correo,
            "Placa": request.form.get("placa", "").upper().strip(),  # ✅ corregido
            "Kilos aproximados": int(request.form.get("kilos")),
            "Pacas": int(request.form.get("pacas"))
        }

        df = pd.concat([df, pd.DataFrame([data])], ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)
        formatear_excel()

        # ── Enviar correo de confirmación (fallo silencioso: no interrumpe el flujo) ──
        enviar_correo_confirmacion(
            destinatario=correo,
            nombre=nombre,
            fecha=fecha,
            cupo=cupo
        )

        return redirect(url_for("sistema", success=1, cupo=cupo, fecha=fecha))

    except PermissionError:
        return jsonify({
            "error": "archivo_bloqueado",
            "mensaje": "El sistema se encuentra actualmente ocupado. Por favor, espere un momento."
        }), 503

    except Exception as e:
        print("ERROR:", e)
        return render_template("error.html", mensaje="Ocurrio un error inesperado. Intente mas tarde."), 500


@app.after_request
def no_cache(response):
    response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, private"
    response.headers["Pragma"] = "no-cache"
    response.headers["Expires"] = "0"
    return response


# ---------- MANEJO GLOBAL DE ERRORES ----------
@app.errorhandler(400)
def error_400(e):
    return render_template("error.html", mensaje="Solicitud invalida"), 400

@app.errorhandler(403)
def error_403(e):
    return render_template("error.html", mensaje="Acceso prohibido"), 403

@app.errorhandler(404)
def error_404(e):
    return render_template("error.html", mensaje="Pagina no encontrada"), 404

@app.errorhandler(409)
def error_409(e):
    return render_template("error.html", mensaje="Conflicto en la solicitud"), 409

@app.errorhandler(500)
def error_500(e):
    return render_template("error.html", mensaje="Error interno del servidor"), 500


# ----------- VALIDACIÓN CENTRAL ---------
def validar_datos(form):
    errores = []

    if not form.get("nombre") or len(form.get("nombre")) > 30:
        errores.append("Nombre invalido")

    if not form.get("proveedor") or len(form.get("proveedor")) > 30:
        errores.append("Proveedor invalido")

    fecha = form.get("fecha")
    if not fecha:
        errores.append("Fecha invalida")
    else:
        try:
            datetime.strptime(fecha, "%Y-%m-%d")
        except Exception:
            errores.append("Fecha invalida")

    cupo = form.get("cupo")
    if not cupo or not re.fullmatch(r"\d+", str(cupo)):
        errores.append("Cupo invalido")
    else:
        try:
            cupo_int = int(cupo)
            if cupo_int < 1 or cupo_int > 100:
                errores.append("Cupo fuera de rango")
        except Exception:
            errores.append("Cupo invalido")

    if not re.fullmatch(r"\d{10}", form.get("telefono", "")):
        errores.append("Telefono invalido")

    telefono_sesion = session.get("usuario", "")
    if not re.fullmatch(r"\d{10}", telefono_sesion):
        errores.append("Telefono de sesion invalido")

    placa = form.get("placa", "").upper().strip()
    if not re.fullmatch(r"[A-Z]{3}[0-9]{3}", placa):
        errores.append("Placa invalida")

    try:
        kilos = int(form.get("kilos"))
        if kilos < 1 or kilos > 50000:
            errores.append("Kilos fuera de rango")
    except:
        errores.append("Kilos invalidos")

    try:
        pacas = int(form.get("pacas"))
        if pacas < 1 or pacas > 80:
            errores.append("Pacas fuera de rango")
    except:
        errores.append("Pacas invalidas")

    return errores


# ---------- EJECUCIÓN ----------
if __name__ == "__main__":
    app.run(
        host="0.0.0.0",
        port=5000,
        debug=True
    )