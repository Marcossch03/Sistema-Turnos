import pandas as pd
from datetime import datetime, timedelta, date
import tkinter as tk
from tkinter import messagebox, simpledialog
from tkcalendar import Calendar
import os

# === Configuración de archivos ===
CLIENTES_FILE = "clientes.xlsx"
TURNOS_FILE   = "turnos.xlsx"

# Crear archivos si no existen
if not os.path.exists(CLIENTES_FILE):
    pd.DataFrame(columns=["DNI", "Nombre", "Teléfono", "Detalle"]).to_excel(CLIENTES_FILE, index=False)
if not os.path.exists(TURNOS_FILE):
    pd.DataFrame(columns=["DNI", "Cliente", "Fecha", "Hora", "Servicio"]).to_excel(TURNOS_FILE, index=False)

# === Generar lista de horarios diarios ===
def generar_horarios():
    return [f"{h:02d}:{m:02d}" for h in range(9, 18) for m in (0, 30)]

# === Determinar color para cada día ===
def obtener_colores_dias(df_turnos):
    colores = {}
    horarios = generar_horarios()
    hoy = datetime.now()
    for i in range(30):
        dia = hoy + timedelta(days=i)
        fecha_str = dia.strftime("%d/%m/%Y")
        turnos_dia = df_turnos[df_turnos["Fecha"] == fecha_str]
        colores[fecha_str] = "orange" if len(turnos_dia) >= len(horarios) else "green"
    return colores

# === Función para pedir datos de cliente y guardar turno ===
def schedule_cliente(fecha, hora):
    df_clientes = pd.read_excel(CLIENTES_FILE)
    df_turnos   = pd.read_excel(TURNOS_FILE)

    dni = simpledialog.askstring("Cliente", "Ingrese el DNI:")
    if not dni or not dni.isdigit():
        messagebox.showwarning("Error", "DNI inválido.")
        return

    cliente = df_clientes[df_clientes["DNI"] == dni]
    if cliente.empty:
        nombre   = simpledialog.askstring("Nuevo Cliente", "Nombre:")
        telefono = simpledialog.askstring("Nuevo Cliente", "Teléfono:")
        detalle  = simpledialog.askstring("Nuevo Cliente", "Detalle (opcional):")
        nuevo    = pd.DataFrame([[dni, nombre, telefono, detalle]],
                                columns=["DNI", "Nombre", "Teléfono", "Detalle"])
        df_clientes = pd.concat([df_clientes, nuevo], ignore_index=True)
        df_clientes.to_excel(CLIENTES_FILE, index=False)
        cliente_nombre = nombre
    else:
        cliente_nombre = cliente.iloc[0]["Nombre"]

    servicio = simpledialog.askstring("Servicio", "Ingrese el servicio a realizar:")
    ya_tiene = df_turnos[
        (df_turnos["DNI"] == dni) &
        (df_turnos["Fecha"] == fecha) &
        (df_turnos["Hora"] == hora)
    ]
    if not ya_tiene.empty:
        messagebox.showwarning("Duplicado", "Este cliente ya tiene un turno en ese horario.")
        return

    nuevo_turno = pd.DataFrame([[dni, cliente_nombre, fecha, hora, servicio]],
                               columns=["DNI", "Cliente", "Fecha", "Hora", "Servicio"])
    df_turnos = pd.concat([df_turnos, nuevo_turno], ignore_index=True)
    df_turnos.to_excel(TURNOS_FILE, index=False)
    messagebox.showinfo("Éxito", "Turno agendado correctamente.")

# === Mostrar horarios de un día seleccionado ===
def mostrar_turnos_por_dia(fecha):
    df_turnos = pd.read_excel(TURNOS_FILE)
    horarios   = generar_horarios()

    ventana = tk.Toplevel()
    ventana.title(f"Turnos para {fecha}")
    ventana.geometry("350x400")

    tk.Label(ventana, text=f"Horarios para {fecha}", font=("Arial", 14)).pack(pady=10)
    lista = tk.Listbox(ventana, width=25, height=15)
    lista.pack(pady=5)

    for hora in horarios:
        estado = "DISPONIBLE" if df_turnos[(df_turnos["Fecha"] == fecha) & (df_turnos["Hora"] == hora)].empty else "OCUPADO"
        lista.insert(tk.END, f"{hora} - {estado}")

    def seleccionar_hora():
        sel = lista.curselection()
        if not sel:
            messagebox.showwarning("Atención", "Seleccioná un horario.")
            return
        texto = lista.get(sel[0])
        hora, estado = texto.split(" - ")
        if estado == "OCUPADO":
            messagebox.showwarning("Ocupado", "Ese horario ya está ocupado.")
            return
        ventana.destroy()
        schedule_cliente(fecha, hora)

    tk.Button(ventana, text="Seleccionar", command=seleccionar_hora).pack(pady=10)

# === Ventana de calendario con colores y selección visible ===
def calendario_view():
    df_turnos = pd.read_excel(TURNOS_FILE)
    colores   = obtener_colores_dias(df_turnos)

    ventana = tk.Toplevel()
    ventana.title("Calendario de Turnos")

    hoy = date.today()
    cal = Calendar(
        ventana,
        selectmode="day",
        date_pattern="dd/mm/yyyy",
        mindate=hoy,
        maxdate=hoy + timedelta(days=50),
        background="white",
        foreground="black",
        headersbackground="lightgray",
        headersforeground="black",
        selectbackground="skyblue",
        selectforeground="white"
    )
    cal.pack(fill="both", expand=True, padx=10, pady=10)

    # Aplicar colores a cada fecha
    for fecha_str, color in colores.items():
        date_obj = datetime.strptime(fecha_str, "%d/%m/%Y").date()
        cal.calevent_create(date_obj, '', tags=fecha_str)
        cal.tag_config(fecha_str, background=color)

    def on_select():
        fecha_str = cal.get_date()
        fecha_sel = datetime.strptime(fecha_str, "%d/%m/%Y").date()
        if fecha_sel < date.today() or fecha_sel > date.today() + timedelta(days=30):
            messagebox.showwarning("Fecha inválida", "Solo podés agendar entre hoy y los próximos 30 días.")
            return
        mostrar_turnos_por_dia(fecha_str)

    tk.Button(ventana, text="Seleccionar Día", command=on_select).pack(pady=10)

# === Interfaz principal ===
def main():
    root = tk.Tk()
    root.title("MIMO COL Beauty - Turnos")
    root.geometry("400x300")

    tk.Label(root, text="MIMO COL Beauty", font=("Arial", 18)).pack(pady=20)
    tk.Button(root, text="Calendario de Turnos", command=calendario_view, width=30, height=2).pack(pady=10)
    tk.Button(root, text="Salir", command=root.destroy, width=30, height=2).pack(pady=10)

    root.mainloop()

if __name__ == "__main__":
    main()
