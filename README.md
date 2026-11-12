# Automatizacion de datos SUNARP

Descarga semanal de los xlsx oficiales de SUNARP (Registro de propiedad inmueble,
publicados en gob.pe) y conversion a JSON local para la app de propiedades transferidas.

## Puesta en marcha
1. Copia `scripts/`, `.github/` y `data/` a la raiz del repo.
2. En GitHub: Settings > Actions > General > Workflow permissions > "Read and write permissions".
3. Pestana Actions > "Actualizar datos SUNARP" > Run workflow.
4. Abre `data/sunarp/_perfil.json`: dice que columnas trae cada archivo y si hay nivel
   distrito, oficina registral, zona registral o departamento. Con eso se define la app.

## Si no detecta los enlaces
Crea `data/sunarp/urls_manual.json`:
{"transferencia_propiedad": "https://...xlsx", "compraventa_predial_sid": "https://...xlsx"}

## Salidas
- `data/sunarp/transferencia_propiedad.json`, `compraventa_predial_sid.json`
- `data/sunarp/meta.json` (texto de pie: "Datos al [fecha], fuente: SUNARP")
- `data/sunarp/raw/*.xlsx` (copia original para auditoria)
