from flask import Flask, render_template, request, redirect, url_for, jsonify, session
from werkzeug.security import generate_password_hash, check_password_hash
from openpyxl import load_workbook
from openpyxl.styles import Border, Side, Font, Alignment
from datetime import datetime, date

import pandas as pd
import os
import re



app = Flask(__name__)

app.secret_key = "clave-secreta-sgf"  # Cambia esto por una clave segura en producción

EXCEL_FILE = "cupos.xlsx"


# ---------- CREAR EXCEL SI NO EXISTE ----------
COLUMNAS = [
    "Fecha",
    "Nombre del conductor",
    "Tipo de vehículo",
    "Cupo",
    "Proveedor",
    "Telefono",
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

    # Encabezados
    for col in range(1, ws.max_column + 1):
        celda = ws.cell(row=1, column=col)
        celda.font = Font(bold=True)
        celda.alignment = Alignment(horizontal="center", vertical="center")
        celda.border = borde_fino
        ws.column_dimensions[celda.column_letter].width = 22

    # Datos
    for row in range(2, ws.max_row + 1):
        for col in range(1, ws.max_column + 1):
            celda = ws.cell(row=row, column=col)
            celda.border = borde_fino
            celda.alignment = Alignment(horizontal="center", vertical="center")

            # Teléfono como texto (columna F)
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

    # Encabezados
    for col in range(1, ws.max_column + 1):
        celda = ws.cell(row=1, column=col)
        celda.font = Font(bold=True)
        celda.alignment = Alignment(horizontal="center", vertical="center")
        celda.border = borde_fino
        ws.column_dimensions[celda.column_letter].width = 22

    # Datos
    for row in range(2, ws.max_row + 1):
        for col in range(1, ws.max_column + 1):
            celda = ws.cell(row=row, column=col)
            celda.border = borde_fino
            celda.alignment = Alignment(horizontal="center", vertical="center")

            # Teléfono como texto (columna B)
            if col == 2 or celda.column_letter == "B":
                celda.number_format = "@"

    wb.save(USUARIOS_FILE)


# ---------- Registro ----------
@app.route("/registro", methods=["GET", "POST"])
def registro():
    if "usuario" in session:
        return redirect(url_for("sistema"))
    
    if request.method == "POST":
        try:
            df = pd.read_excel(USUARIOS_FILE)

            # Limpiar y normalizar teléfono recibido
            telefono = request.form.get("telefono", "").strip()
            telefono = re.sub(r"\D", "", telefono)

            # Normalizar columna existente en el archivo (evita .0 al leer desde Excel)
            if "Telefono" in df.columns:
                df["Telefono"] = df["Telefono"].astype(str).str.replace(r"\.0$", "", regex=True).str.strip()
            else:
                df["Telefono"] = df.get("Telefono", pd.Series(dtype=str))

            if telefono in df["Telefono"].values:
                return render_template("error.html",
                    mensaje="El teléfono ya está registrado"
                )

            nuevo = {
                "Nombre": request.form.get("nombre"),
                "Telefono": telefono,
                "Proveedor": request.form.get("proveedor"),
                "Password": generate_password_hash(request.form.get("password"))
            }

            df = pd.concat([df, pd.DataFrame([nuevo])], ignore_index=True)
            df.to_excel(USUARIOS_FILE, index=False)
            formatear_excel_usuarios()

            # Guardar teléfono normalizado en sesión
            session["usuario"] = telefono
            return redirect(url_for("sistema"))

        except PermissionError:
            return jsonify({
                "error": "archivo_bloqueado",
                "mensaje": "El sistema se encuentra actualmente ocupado. Por favor, espere un momento."
            }), 503

        except Exception as e:
            print("ERROR:", e)
            return render_template("error.html",
                mensaje="Ocurrió un error inesperado. Intente más tarde."
            ), 500

    
    return render_template("formulario.html")


@app.route("/login", methods=["GET", "POST"])
def login():
    if "usuario" in session:
        return redirect(url_for("sistema"))
    
    if request.method == "POST":
        # Limpiar/normalizar datos de entrada
        telefono = request.form.get("telefono", "").strip()
        telefono = re.sub(r"\D", "", telefono)
        password = request.form.get("password", "")

        df = pd.read_excel(USUARIOS_FILE)
        # Normalizar columna Telefono para comparaciones seguras
        if "Telefono" in df.columns:
            df["Telefono"] = df["Telefono"].astype(str).str.replace(r"\.0$", "", regex=True).str.strip()

        user = df[df["Telefono"] == telefono]

        if user.empty:
            return render_template("error.html", mensaje="Usuario no encontrado")

        if not check_password_hash(user.iloc[0]["Password"], password):
            return render_template("error.html", mensaje="Contraseña incorrecta")

        session["usuario"] = telefono
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
    return render_template("sistema.html")

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


@app.route("/guardar", methods=["POST"])
def guardar_cupo():
    if "usuario" not in session:
        return redirect(url_for("login"))
    try:
        asegurar_excel()

        errores = validar_datos(request.form)
        if errores:
            return render_template("error.html",
                mensaje="Datos inválidos enviados al sistema"
            ), 400

        fecha = request.form.get("fecha")
        cupo = int(request.form.get("cupo"))

        # ---------- VALIDACIONES DE FECHA EN BACKEND ----------
        fecha_cupo = datetime.strptime(fecha, "%Y-%m-%d").date()
        hoy = date.today()

        # Fechas pasadas
        if fecha_cupo < hoy:
            return render_template("error.html",
                mensaje="No se pueden registrar cupos para fechas pasadas"
            ), 400

        # Domingos
        if fecha_cupo.weekday() == 6:
            return render_template("error.html",
                mensaje="No se pueden registrar cupos los domingos"
            ), 400

        # Límite de cupos por día
        if fecha_cupo.weekday() <= 4:
            max_cupos = 12
        else:
            max_cupos = 5

        if cupo < 1 or cupo > max_cupos:
            return render_template("error.html",
                mensaje="El número de cupo no es válido para ese día"
            ), 400

        # ---------- VALIDACIÓN CONTRA EXCEL ----------
        df = pd.read_excel(EXCEL_FILE)

        if ((df["Fecha"] == fecha) & (df["Cupo"] == cupo)).any():
            return render_template("error.html",
                mensaje="El cupo seleccionado ya está ocupado"
            ), 409

        # ---------- GUARDAR ----------
        data = {
            "Fecha": fecha,
            "Nombre del conductor": request.form.get("nombre"),
            "Tipo de vehículo": request.form.get("tipo"),
            "Cupo": cupo,
            "Proveedor": request.form.get("proveedor"),
            "Telefono": str(request.form.get("telefono")),
            "Placa": session.get("usuario"),
            "Kilos aproximados": int(request.form.get("kilos")),
            "Pacas": int(request.form.get("pacas"))
        }

        df = pd.concat([df, pd.DataFrame([data])], ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)
        formatear_excel()

        return redirect(url_for("sistema", success=1))

    except PermissionError:
        return jsonify({
            "error": "archivo_bloqueado",
            "mensaje": "El sistema se encuentra actualmente ocupado. Por favor, espere un momento."
        }), 503

    except Exception as e:
        print("ERROR:", e)
        return render_template("error.html",
            mensaje="Ocurrió un error inesperado. Intente más tarde."
        ), 500
    

@app.after_request
def no_cache(response):
    response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, private"
    response.headers["Pragma"] = "no-cache"
    response.headers["Expires"] = "0"
    return response



# ---------- MANEJO GLOBAL DE ERRORES ----------
@app.errorhandler(400)
def error_400(e):
    return render_template("error.html", mensaje="Solicitud inválida"), 400

@app.errorhandler(403)
def error_403(e):
    return render_template("error.html", mensaje="Acceso prohibido"), 403

@app.errorhandler(404)
def error_404(e):
    return render_template("error.html", mensaje="Página no encontrada"), 404

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
        errores.append("Nombre inválido")

    if not form.get("proveedor") or len(form.get("proveedor")) > 30:
        errores.append("Proveedor inválido")

    # Fecha: required and must be YYYY-MM-DD
    fecha = form.get("fecha")
    if not fecha:
        errores.append("Fecha inválida")
    else:
        try:
            datetime.strptime(fecha, "%Y-%m-%d")
        except Exception:
            errores.append("Fecha inválida")

    # Cupo: required, must be a positive integer
    cupo = form.get("cupo")
    if not cupo or not re.fullmatch(r"\d+", str(cupo)):
        errores.append("Cupo inválido")
    else:
        try:
            cupo_int = int(cupo)
            if cupo_int < 1 or cupo_int > 100:
                errores.append("Cupo fuera de rango")
        except Exception:
            errores.append("Cupo inválido")

    if not re.fullmatch(r"\d{10}", form.get("telefono", "")):
        errores.append("Teléfono inválido")

    telefono_sesion = session.get("usuario", "")
    if not re.fullmatch(r"\d{10}", telefono_sesion):
        errores.append("Teléfono inválido")

    try:
        kilos = int(form.get("kilos"))
        if kilos < 1 or kilos > 50000:
            errores.append("Kilos fuera de rango")
    except:
        errores.append("Kilos inválidos")

    try:
        pacas = int(form.get("pacas"))
        if pacas < 1 or pacas > 80:
            errores.append("Pacas fuera de rango")
    except:
        errores.append("Pacas inválidas")

    return errores


# ---------- EJECUCIÓN ----------
if __name__ == "__main__":
    app.run(
        host="0.0.0.0",
        port=5000,
        debug=True
    )