# Generador de Nombres de Productos

Sistema automatizado para generar nombres descriptivos de productos a partir de datos almacenados en MongoDB.

## Descripción

El proyecto genera nombres de productos estructurados consultando una base de datos MongoDB y aplicando transformaciones según esquemas configurables. El sistema procesa productos en diferentes etapas de enriquecimiento de datos:

- **Productos originales**: Estado inicial de los productos
- **Productos después de Abasteo**: Datos enriquecidos mediante el sistema Abasteo
- **Productos después de Icecat**: Datos complementados con información de Icecat
- **Productos completos**: Estado final con toda la información disponible

## Características

- Generación automática de nombres de productos basados en especificaciones técnicas
- Soporte para múltiples categorías (laptops, monitores, impresoras, mouse gamer, etc.)
- Exportación a archivos Excel con pestañas diferenciadas por colores
- Transformaciones de texto (mayúsculas, minúsculas, singularización)
- Detección y marcado de datos faltantes

## Tecnologías

- Python
- MongoDB (PyMongo)
- OpenPyXL (generación de Excel)
- Icecat (enriquecimiento de datos)
- Abasteo (enriquecimiento de datos)

