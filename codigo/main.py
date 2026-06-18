from openpyxl import load_workbook
from datetime import datetime


ARCHIVO_EXCEL = "../base_datos/empleados.xlsx"


def buscar_empleado(legajo):
    wb = load_workbook(ARCHIVO_EXCEL)
    hoja = wb.active

    for fila in hoja.iter_rows(min_row=2):

        if str(fila[0].value) == legajo:

            return {
                "fila": fila[0].row,
                "nombre": fila[1].value,
                "sector": fila[2].value,
                "dias": int(fila[3].value),
                "workbook": wb,
                "hoja": hoja
            }

    wb.close()
    return None


def validar_fecha(fecha):

    try:

        fecha_convertida = datetime.strptime(
            fecha,
            "%d/%m/%Y"
        )

        if fecha_convertida.date() < datetime.today().date():
            return False

        return True

    except ValueError:
        return False


def solicitar_legajo():

    while True:

        legajo = input(
            "\nIngrese su legajo: "
        )

        empleado = buscar_empleado(legajo)

        if empleado:
            return empleado

        print("Error: legajo inexistente.")


def solicitar_fecha():

    while True:

        fecha = input(
            "\nIngrese fecha de inicio (dd/mm/aaaa): "
        )

        if validar_fecha(fecha):
            return fecha

        print("Error: fecha inválida.")


def solicitar_dias():

    while True:

        try:

            dias = int(
                input(
                    "\nIngrese cantidad de días solicitados: "
                )
            )

            if dias <= 0:

                print(
                    "Error: la cantidad debe ser mayor a cero."
                )

                continue

            return dias

        except ValueError:

            print(
                "Error: debe ingresar un número entero."
            )


def solicitar_aprobacion_supervisor():

    while True:

        respuesta = input(
            "\n¿Supervisor aprueba la solicitud? (S/N): "
        ).upper()

        if respuesta in ["S", "N"]:
            return respuesta

        print(
            "Error: ingrese S para aprobar o N para rechazar."
        )


def actualizar_saldo(empleado, dias_solicitados):

    saldo_actual = empleado["dias"] - dias_solicitados

    empleado["hoja"].cell(
        row=empleado["fila"],
        column=4
    ).value = saldo_actual

    empleado["workbook"].save(
        ARCHIVO_EXCEL
    )

    empleado["workbook"].close()

    return saldo_actual


def generar_comprobante(
    empleado,
    fecha,
    dias,
    saldo_anterior,
    saldo_actual
):

    print("\n")
    print("=" * 50)
    print("COMPROBANTE DE SOLICITUD DE VACACIONES")
    print("=" * 50)

    print(f"Empleado: {empleado['nombre']}")
    print(f"Sector: {empleado['sector']}")
    print(f"Fecha de inicio: {fecha}")
    print(f"Días solicitados: {dias}")
    print(f"Saldo anterior: {saldo_anterior}")
    print(f"Saldo actual: {saldo_actual}")

    print("=" * 50)


def main():

    print("=" * 50)
    print("BCH SISTEMAS")
    print("GESTIÓN DE VACACIONES")
    print("=" * 50)

    # ----------------------------------
    # Solicitar legajo
    # ----------------------------------

    empleado = solicitar_legajo()

    # ----------------------------------
    # Mostrar información del empleado
    # ----------------------------------

    print("\nInformación del empleado")

    print(f"Nombre: {empleado['nombre']}")
    print(f"Sector: {empleado['sector']}")
    print(
        f"Días disponibles: {empleado['dias']}"
    )

    # ----------------------------------
    # Solicitar fecha
    # ----------------------------------

    fecha = solicitar_fecha()

    # ----------------------------------
    # Solicitar cantidad de días
    # ----------------------------------

    dias_solicitados = solicitar_dias()

    # ----------------------------------
    # Consultar saldo disponible
    # ----------------------------------

    if dias_solicitados > empleado["dias"]:

        print("\nSolicitud rechazada.")
        print(
            f"Días disponibles: {empleado['dias']}"
        )

        empleado["workbook"].close()
        return

    # ----------------------------------
    # Evaluar aprobación
    # ----------------------------------

    if dias_solicitados > 10:

        print(
            "\nLa solicitud debe ser revisada "
            "por un supervisor."
        )

        aprobacion = (
            solicitar_aprobacion_supervisor()
        )

        if aprobacion == "N":

            print(
                "\nSolicitud rechazada por el supervisor."
            )

            empleado["workbook"].close()
            return

        print(
            "\nSolicitud aprobada por el supervisor."
        )

    else:

        print(
            "\nSolicitud aprobada automáticamente."
        )

    # ----------------------------------
    # Actualizar base de datos
    # ----------------------------------

    saldo_anterior = empleado["dias"]

    saldo_actual = actualizar_saldo(
        empleado,
        dias_solicitados
    )

    # ----------------------------------
    # Generar comprobante
    # ----------------------------------

    generar_comprobante(
        empleado,
        fecha,
        dias_solicitados,
        saldo_anterior,
        saldo_actual
    )

    # ----------------------------------
    # Confirmación final
    # ----------------------------------

    print(
        "\nLa solicitud fue registrada correctamente."
    )

    print(
        "Confirmación enviada al empleado."
    )

    print(
        "\nProceso finalizado."
    )


if __name__ == "__main__":
    main()