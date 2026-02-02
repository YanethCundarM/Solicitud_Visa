# Script para generar archivo Excel en formato abierto
# Uso alternativo sin dependencias de openpyxl

archivo_csv = """FORMULARIO DS-160 - SOLICITUD DE VISA AMERICANA
Para Ciudadanos Colombianos
Generado: 2026-02-01

INSTRUCCIONES IMPORTANTES:
- Este es un formulario de AYUDA para planificar tu solicitud
- DEBES llenar el DS-160 oficial en: https://ceac.state.gov/GenNIV/
- Toda la información debe ser HONESTA y VERIFICABLE
- Mentir es FRAUDE y tiene consecuencias legales

DOCUMENTOS REQUERIDOS:
✓ Pasaporte válido (vigencia 6+ meses)
✓ Cédula de ciudadanía
✓ Foto tamaño 5x5 cm (fondo blanco)
✓ Comprobante de pago ($160 USD)
✓ Extractos bancarios (3 últimos meses)
✓ Carta del empleador o comprobante de ingresos
✓ Comprobante de vivienda
✓ Si viaja por negocios: Carta de empleador en EE.UU.

================================================================================
SECCIÓN 1: INFORMACIÓN PERSONAL BÁSICA
================================================================================

Apellido;
Nombre;
Segundo Nombre;
Otros Apellidos;
Fecha de Nacimiento (DD/MM/YYYY);
Lugar de Nacimiento;
Género (M/F);
Nacionalidad;Colombiano/a
Número de Cédula;
Número de Pasaporte;

CONTACTO:;
Correo Electrónico;
Teléfono Principal;
Teléfono Secundario;

================================================================================
SECCIÓN 2: INFORMACIÓN DE PASAPORTE
================================================================================

Número de Pasaporte;
País de Emisión;Colombia
Fecha de Emisión (DD/MM/YYYY);
Fecha de Vencimiento (DD/MM/YYYY);
Observaciones;Debe ser válido 6+ meses desde la solicitud

================================================================================
SECCIÓN 3: ANTECEDENTES Y VIAJES
================================================================================

¿Ha viajado a EE.UU. antes? (Sí/No);
Año del último viaje;
Propósito del viaje anterior;
¿Tiene antecedentes penales? (Sí/No);No
Descripción de antecedentes;
¿Visas anteriores de EE.UU.? (Sí/No);
Tipos de visas anteriores;

================================================================================
SECCIÓN 4: INFORMACIÓN DEL VIAJE PLANEADO
================================================================================

Tipo de Visa (B1/B2/F1/H1B/L1);
Fecha de Llegada Planeada;
Ciudad/Estado de Destino;
Duración del Viaje;
Propósito Principal del Viaje;

CONTACTO EN EE.UU.:;
Empresa/Institución;
Dirección Completa;
Ciudad y Estado;
Teléfono;
Correo;

================================================================================
SECCIÓN 5: INFORMACIÓN LABORAL Y ECONÓMICA
================================================================================

Estatus de Empleo;
Ocupación/Cargo;
Nombre del Empleador;
Dirección del Empleador;
Años en el Empleo Actual;
Ingresos Anuales Aproximados (COP);
¿Quién financia el viaje?;

REFERENCIAS:;
Nombre del Supervisor;
Teléfono;
Correo;

================================================================================
SECCIÓN 6: INFORMACIÓN FAMILIAR
================================================================================

PADRE:;
Nombre Completo;
Fecha de Nacimiento;
¿Aún vive? (Sí/No);

MADRE:;
Nombre Completo;
Fecha de Nacimiento;
¿Aún vive? (Sí/No);

HERMANOS:;
¿Tienes hermanos? ¿Cuántos?;
¿Alguno vive en EE.UU.?;Especificar nombres y años

ESTADO CIVIL:;
Estado Civil;
¿Tienes hijos? ¿Cuántos?;
Edades de los hijos;

================================================================================
SECCIÓN 7: DOCUMENTOS Y DECLARACIONES
================================================================================

CHECKLIST DE DOCUMENTOS:;
☐ Pasaporte original;
☐ Cédula original;
☐ Foto tamaño 5x5 cm;
☐ Confirmación de pago DS-160;
☐ Extractos bancarios;
☐ Carta de empleador;
☐ Comprobante de vivienda;
☐ Comprobante de fondos;

DECLARACIONES LEGALES:;
☐ Declaro que toda la información es verdadera;
☐ Entiendo que mentir es FRAUDE;
☐ Acepto todos los términos y condiciones;

================================================================================
PRÓXIMOS PASOS
================================================================================

1. COMPLETA este formulario como guía
2. LLENA el DS-160 oficial en ceac.state.gov
3. PAGA la tarifa de $160 USD
4. PROGRAMA tu cita en ais.usvisa-info.com/es-co/
5. RECOPILA todos los documentos
6. ASISTE a tu entrevista

CONTACTOS ÚTILES:
- Embajada de EE.UU. en Colombia: https://co.usembassy.gov/
- Sistema de Citas: https://ais.usvisa-info.com/es-co/
- Información de Visas: https://travel.state.gov/
- Formulario DS-160: https://ceac.state.gov/GenNIV/

IMPORTANTE: Esta es solo una herramienta de ayuda. El formulario oficial 
debe completarse en el sitio de CEAC (Departamento de Estado de EE.UU.)
"""

# Guardar como CSV que Excel puede abrir
with open("Formulario_Visa_DS160_Guia.csv", "w", encoding="utf-8") as f:
    f.write(archivo_csv)

print("✅ Archivo CSV creado exitosamente: Formulario_Visa_DS160_Guia.csv")
print("   Este archivo puede abrirse con Excel o Google Sheets")
