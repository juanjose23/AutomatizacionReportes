"# AutomatizacionReportes" 
# Reporte General

Este proyecto realiza la comparación de inventarios entre datos de compra y venta, y genera un reporte en formato Excel con información adicional sobre precios y materiales.

## Descripción

El script en Python realiza las siguientes tareas:
1. Importa datos de archivos de compra y venta.
2. Combina y compara los datos para calcular inventarios teóricos y físicos.
3. Importa precios y clasificaciones de materiales desde un archivo de Excel.
4. Calcula la diferencia entre inventarios y el valor monetario asociado.
5. Guarda los resultados en un archivo Excel con formato y estilo corporativo.

## Requisitos

Asegúrate de tener las siguientes dependencias instaladas:

- `pandas`: Para manipulación de datos.
- `openpyxl`: Para manipulación de archivos Excel.

Puedes instalar estas dependencias utilizando el archivo `requirements.txt`.

## Instalación

1. Clona el repositorio:
   ```bash
   git clone https://github.com/juanjose23/AutomatizacionReportes.git

2. Navega al directorio del proyecto:
   ```bash
   cd AutomatizacionReportes

3. Crea un entorno virtual (opcional pero recomendado):
   ```bash
   python -m venv env

4. Activa el entorno virtual desde el CMD:
   ```bash
   .\env\Scripts\activate

5. Instala las dependencias:
   ```bash
   pip install -r requirements.txt

6. USO:
   ```bash
   Ejecuta el script principal con:
   python ReporteGeneral.py




