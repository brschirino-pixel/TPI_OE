from openpyxl import load_workbook
from datetime import datetime

# ============================================================
# CONFIGURACIÓN DE ARCHIVOS
# ============================================================
ARCHIVO_EMPLEADOS = "../base_datos/empleados.xlsx"
ARCHIVO_SOLICITUDES = "../base_datos/solicitudes.xlsx"


# ============================================================
# FUNCIONES DE ACCESO A DATOS
# ============================================================

def buscar_empleado(legajo):
    """
    Busca un empleado en el archivo empleados.xlsx por su número de legajo.
    
    Args:
        legajo (str): Número de legajo del empleado
        
    Returns:
        dict: Diccionario con los datos del empleado o None si no existe
    """
    # Abrir el archivo de empleados
    wb = load_workbook(ARCHIVO_EMPLEADOS)
    hoja = wb.active

    # Recorrer todas las filas (desde la 2 para saltar el encabezado)
    for fila in hoja.iter_rows(min_row=2):

        # Verificar si el legajo coincide
        if str(fila[0].value) == legajo:

            # Retornar todos los datos del empleado
            return {
                "fila": fila[0].row,           # Número de fila para actualizar después
                "legajo": fila[0].value,       # Número de legajo
                "nombre": fila[1].value,       # Nombre completo
                "sector": fila[2].value,       # Sector donde trabaja
                "dias": int(fila[3].value),    # Días disponibles
                "workbook": wb,                # Objeto del libro para guardar cambios
                "hoja": hoja                   # Hoja activa para actualizar
            }

    # Si no se encontró, cerrar el archivo y retornar None
    wb.close()
    return None

# ============================================================
# FUNCIONES DE VALIDACIÓN
# ============================================================

def validar_fecha(fecha):
    """
    Valida que una fecha tenga formato dd/mm/aaaa y sea posterior a hoy.
    
    Args:
        fecha (str): Fecha en formato dd/mm/aaaa
        
    Returns:
        tuple: (bool, str/mensaje) - (True, fecha_convertida) si es válida, (False, mensaje_error) si no
    """
    try:
        # Intentar convertir la fecha al formato dd/mm/aaaa
        fecha_convertida = datetime.strptime(fecha, "%d/%m/%Y")

        # Verificar que la fecha sea posterior a hoy
        if fecha_convertida.date() < datetime.today().date():
            return False, "La fecha debe ser posterior a hoy."

        return True, fecha_convertida

    except ValueError:
        # Si el formato es incorrecto, mostrar mensaje de error
        return False, "Formato incorrecto. Use dd/mm/aaaa."


def validar_legajo(legajo):
    """
    Valida que el legajo sea solo números, sin espacios ni caracteres especiales.
    
    Args:
        legajo (str): Número de legajo ingresado por el usuario
        
    Returns:
        tuple: (bool, str) - (True, legajo_limpio) si es válido, (False, mensaje_error) si no
    """
    # Eliminar espacios en blanco al inicio y final
    legajo = legajo.strip()
    
    # Verificar que no esté vacío
    if not legajo:
        return False, "El legajo no puede estar vacío."
    
    # Verificar que solo contenga números
    if not legajo.isdigit():
        return False, "El legajo debe contener solo números."
    
    return True, legajo


# ============================================================
# FUNCIONES DE INTERACCIÓN CON EL USUARIO (SOLICITAR DATOS)
# ============================================================

def solicitar_legajo():
    """
    Solicita al usuario que ingrese su número de legajo.
    Valida que exista en la base de datos y que tenga formato correcto.
    
    Returns:
        dict: Datos del empleado encontrado
    """
    while True:
        # Mostrar mensaje al usuario
        print("\nBOT: Para comenzar, ingrese su número de legajo:")
        legajo_input = input("Legajo: ").strip()

        # Verificar si el usuario quiere salir
        if legajo_input.lower() == "salir":
            print("\nBOT: Conversación finalizada.")
            exit()

        # Validar el formato del legajo
        es_valido, mensaje = validar_legajo(legajo_input)
        if not es_valido:
            print(f"BOT: {mensaje}")
            continue  # Volver a pedir el legajo

        # Buscar el empleado en la base de datos
        empleado = buscar_empleado(legajo_input)

        if empleado:
            return empleado  # Empleado encontrado, continuar

        # Si no se encontró, mostrar error y volver a pedir
        print("BOT: No se encontró un empleado con ese legajo. Verifique e intente nuevamente.")


def solicitar_fecha():
    """
    Solicita al usuario que ingrese la fecha de inicio de vacaciones.
    Valida formato y que sea una fecha futura.
    
    Returns:
        str: Fecha en formato dd/mm/aaaa
    """
    while True:
        # Mostrar mensaje al usuario con el formato esperado
        print("\nBOT: Ingrese la fecha de inicio (formato dd/mm/aaaa):")
        fecha_input = input("Fecha: ").strip()

        # Verificar si el usuario quiere salir
        if fecha_input.lower() == "salir":
            print("\nBOT: Conversación finalizada.")
            exit()

        # Validar la fecha ingresada
        es_valida, mensaje = validar_fecha(fecha_input)
        if es_valida:
            return fecha_input  # Fecha válida, continuar

        # Si la fecha no es válida, mostrar el error y volver a pedir
        print(f"BOT: {mensaje}")


def solicitar_dias():
    """
    Solicita al usuario la cantidad de días de vacaciones.
    Valida que sea un número entero positivo y no mayor a 30.
    
    Returns:
        int: Cantidad de días solicitados
    """
    while True:
        # Mostrar mensaje al usuario
        print("\nBOT: ¿Cuántos días de vacaciones desea solicitar?")
        dias_input = input("Cantidad de días: ").strip()

        # Verificar si el usuario quiere salir
        if dias_input.lower() == "salir":
            print("\nBOT: Conversación finalizada.")
            exit()

        # Validar que el campo no esté vacío
        if not dias_input:
            print("BOT: Debe ingresar un número. El campo no puede estar vacío.")
            continue

        # Validar que solo contenga números
        if not dias_input.isdigit():
            print("BOT: Debe ingresar un número entero. Ejemplo: 5")
            continue

        # Convertir a número entero
        dias = int(dias_input)

        # Validar que sea mayor a cero
        if dias <= 0:
            print("BOT: La cantidad de días debe ser mayor a cero.")
            continue

        # Validar que no exceda el límite máximo de 30 días
        if dias > 30:
            print("BOT: La cantidad máxima de días que se puede solicitar es 30.")
            continue

        return dias  # Días válidos, continuar


def solicitar_aprobacion_supervisor():
    """
    Solicita al supervisor que apruebe o rechace la solicitud.
    Solo acepta 'S' (aprobado) o 'N' (rechazado).
    
    Returns:
        str: 'S' para aprobado o 'N' para rechazado
    """
    while True:
        # Mostrar mensaje al supervisor
        print("\nBOT: Supervisor, ingrese 'S' para aprobar o 'N' para rechazar:")
        respuesta = input("Aprobación (S/N): ").strip().upper()

        # Procesar la respuesta del supervisor
        if respuesta == "S":
            return "S"  # Solicitud aprobada
        elif respuesta == "N":
            return "N"  # Solicitud rechazada
        elif respuesta.lower() == "salir":
            # Permitir que el supervisor también pueda salir
            print("\nBOT: Conversación finalizada. Gracias por utilizar BCH Sistemas.")
            exit()
        else:
            # Si la respuesta no es válida, mostrar mensaje y volver a pedir
            print("BOT: Respuesta inválida. Ingrese 'S' para aprobar o 'N' para rechazar.")


# ============================================================
# FUNCIONES DE PROCESAMIENTO DE DATOS
# ============================================================

def actualizar_saldo(empleado, dias_solicitados):
    """
    Actualiza los días disponibles del empleado después de aprobar la solicitud.
    
    Args:
        empleado (dict): Datos del empleado
        dias_solicitados (int): Cantidad de días solicitados
        
    Returns:
        int: Nuevo saldo de días disponibles
    """
    # Calcular el nuevo saldo (días disponibles - días solicitados)
    saldo_actual = empleado["dias"] - dias_solicitados

    # Actualizar la celda en Excel
    empleado["hoja"].cell(
        row=empleado["fila"],
        column=4  # Columna D = Días disponibles
    ).value = saldo_actual

    # Guardar los cambios en el archivo
    empleado["workbook"].save(ARCHIVO_EMPLEADOS)
    empleado["workbook"].close()

    return saldo_actual


def registrar_solicitud(empleado, fecha, dias, estado):
    """
    Registra la solicitud en el archivo solicitudes.xlsx para mantener el historial.
    
    Args:
        empleado (dict): Datos del empleado
        fecha (str): Fecha de inicio de vacaciones
        dias (int): Cantidad de días solicitados
        estado (str): 'Aprobada' o 'Rechazada'
    """
    # Abrir el archivo de solicitudes
    wb = load_workbook(ARCHIVO_SOLICITUDES)
    hoja = wb.active

    # Generar un ID único para la solicitud
    nuevo_id = hoja.max_row

    # Agregar una nueva fila con los datos de la solicitud
    hoja.append([
        nuevo_id,                      # ID de la solicitud
        empleado["legajo"],            # Legajo del empleado
        empleado["nombre"],            # Nombre del empleado
        empleado["sector"],            # Sector del empleado
        fecha,                         # Fecha de inicio
        dias,                          # Días solicitados
        estado                         # Estado de la solicitud
    ])

    # Guardar los cambios en el archivo
    wb.save(ARCHIVO_SOLICITUDES)
    wb.close()


def generar_comprobante(empleado, fecha, dias, saldo_anterior, saldo_actual, estado):
    """
    Genera un comprobante impreso con todos los detalles de la solicitud.
    
    Args:
        empleado (dict): Datos del empleado
        fecha (str): Fecha de inicio de vacaciones
        dias (int): Cantidad de días solicitados
        saldo_anterior (int): Días disponibles antes de la solicitud
        saldo_actual (int): Días disponibles después de la solicitud
        estado (str): 'Aprobada' o 'Rechazada'
    """
    # Encabezado del comprobante
    print("\n" + "=" * 60)
    print("COMPROBANTE DE SOLICITUD DE VACACIONES")
    print("=" * 60)

    # Información del empleado y la solicitud
    print(f"Legajo: {empleado['legajo']}")
    print(f"Empleado: {empleado['nombre']}")
    print(f"Sector: {empleado['sector']}")
    print(f"Fecha de inicio: {fecha}")
    print(f"Días solicitados: {dias}")
    print(f"Estado: {estado}")

    # Si la solicitud fue aprobada, mostrar el saldo actualizado
    if estado == "Aprobada":
        print(f"Días disponibles anteriores: {saldo_anterior}")
        print(f"Días disponibles actuales: {saldo_actual}")

    # Pie del comprobante
    print("=" * 60)


# ============================================================
# FUNCIÓN PRINCIPAL (FLUJO DEL CHATBOT)
# ============================================================

def main():
    """
    Función principal que ejecuta el flujo completo del chatbot.
    Gestiona los estados del proceso y la lógica de negocio.
    """
    # Inicializar el estado del proceso
    estado = "INICIO"

    # ============================================================
    # PRESENTACIÓN DEL SISTEMA
    # ============================================================
    print("=" * 60)
    print("CHATBOT DE GESTIÓN DE VACACIONES")
    print("BCH SISTEMAS")
    print("=" * 60)

    print("\nBOT: Bienvenido al sistema de gestión de vacaciones.")
    print("BOT: Este asistente le ayudará a registrar su solicitud de vacaciones.")
    print("BOT: En cualquier momento, escriba 'salir' para finalizar la conversación.")

    # ============================================================
    # ESTADO 1: SOLICITAR LEGAJO
    # ============================================================
    estado = "SOLICITANDO_LEGAJO"
    empleado = solicitar_legajo()

    # ============================================================
    # ESTADO 2: CONSULTAR EMPLEADO
    # ============================================================
    estado = "CONSULTANDO_EMPLEADO"
    print("\nBOT: Empleado encontrado.")

    # Mostrar los datos del empleado
    print(f"BOT: Nombre: {empleado['nombre']}")
    print(f"BOT: Sector: {empleado['sector']}")
    print(f"BOT: Días disponibles: {empleado['dias']}")

    # Verificar si el empleado tiene días disponibles
    if empleado["dias"] == 0:
        print("\nBOT: Usted no tiene días disponibles para solicitar vacaciones.")
        print("BOT: Proceso finalizado.")
        empleado["workbook"].close()
        return

    # ============================================================
    # ESTADO 3: SOLICITAR FECHA
    # ============================================================
    estado = "SOLICITANDO_FECHA"
    fecha = solicitar_fecha()

    # ============================================================
    # ESTADO 4: SOLICITAR DÍAS
    # ============================================================
    estado = "SOLICITANDO_DIAS"
    dias_solicitados = solicitar_dias()

    # ============================================================
    # ESTADO 5: EVALUAR SOLICITUD
    # ============================================================
    estado = "EVALUANDO_SOLICITUD"
    print("\nBOT: Procesando su solicitud...")

    # Verificar si el empleado tiene suficientes días disponibles
    if dias_solicitados > empleado["dias"]:
        # CASO: Días insuficientes → Solicitud rechazada automáticamente
        registrar_solicitud(
            empleado,
            fecha,
            dias_solicitados,
            "Rechazada"
        )

        generar_comprobante(
            empleado,
            fecha,
            dias_solicitados,
            empleado["dias"],
            empleado["dias"],
            "Rechazada"
        )

        print(
            f"\nBOT: Solicitud rechazada. No tiene suficientes días disponibles. (Días disponibles: {empleado['dias']})"
        )

        estado = "FINALIZADO"
        print(f"BOT: Estado actual: {estado}")

        empleado["workbook"].close()
        return

    # ============================================================
    # ESTADO 6: APROBACIÓN POR SUPERVISOR (si aplica)
    # ============================================================
    # Regla de negocio: solicitudes de más de 10 días requieren aprobación del supervisor
    if dias_solicitados > 10:
        estado = "APROBACION_SUPERVISOR"

        print(
            f"\nBOT: Su solicitud de {dias_solicitados} días requiere aprobación del supervisor."
        )

        # Solicitar aprobación al supervisor
        respuesta = solicitar_aprobacion_supervisor()

        if respuesta == "N":
            # CASO: Supervisor rechaza la solicitud
            registrar_solicitud(
                empleado,
                fecha,
                dias_solicitados,
                "Rechazada"
            )

            generar_comprobante(
                empleado,
                fecha,
                dias_solicitados,
                empleado["dias"],
                empleado["dias"],
                "Rechazada"
            )

            print("\nBOT: Solicitud rechazada por el supervisor.")

            estado = "FINALIZADO"
            print(f"BOT: Estado actual: {estado}")

            empleado["workbook"].close()
            return

        # CASO: Supervisor aprueba la solicitud
        print("\nBOT: Solicitud aprobada por el supervisor.")

    else:
        # CASO: Solicitud de hasta 10 días → Aprobación automática
        print("\nBOT: Solicitud aprobada automáticamente (hasta 10 días).")

    # ============================================================
    # ESTADO 7: ACTUALIZAR DATOS
    # ============================================================
    estado = "ACTUALIZANDO_DATOS"

    # Guardar el saldo anterior para el comprobante
    saldo_anterior = empleado["dias"]

    # Actualizar los días disponibles del empleado
    saldo_actual = actualizar_saldo(
        empleado,
        dias_solicitados
    )

    # Registrar la solicitud en el historial
    registrar_solicitud(
        empleado,
        fecha,
        dias_solicitados,
        "Aprobada"
    )

    # ============================================================
    # ESTADO 8: GENERAR COMPROBANTE
    # ============================================================
    estado = "GENERANDO_COMPROBANTE"

    generar_comprobante(
        empleado,
        fecha,
        dias_solicitados,
        saldo_anterior,
        saldo_actual,
        "Aprobada"
    )

    # ============================================================
    # ESTADO FINAL: FINALIZADO
    # ============================================================
    estado = "FINALIZADO"

    print("\nBOT: ¡Solicitud registrada exitosamente!")
    print("BOT: Se ha enviado la confirmación al empleado.")
    print(f"BOT: Estado actual: {estado}")
    print("\nBOT: Gracias por utilizar el sistema de BCH Sistemas.")


# ============================================================
# PUNTO DE ENTRADA DEL PROGRAMA
# ============================================================
if __name__ == "__main__":
    main()