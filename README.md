# Demo — Generador de Convenios Talento EAFIT

## Cómo ejecutar

### 1. Instalar dependencias (solo la primera vez)
```bash
pip install -r requirements.txt
```

### 2. Correr la app
```bash
streamlit run app.py
```

La app se abre automáticamente en el navegador en http://localhost:8501

## Estructura
```
demo/
├── app.py              ← Interfaz (formulario web)
├── generador.py        ← Motor Python (rellena los documentos)
├── crear_plantillas.py ← Script para preparar las plantillas
├── requirements.txt
└── templates/          ← Plantillas Word estandarizadas
    ├── convenio_01_remunerado_org_arl.docx
    ├── convenio_02_remunerado_eafit_arl.docx
    ├── convenio_03_no_remunerado_org_arl.docx
    └── convenio_04_no_remunerado_eafit_arl.docx
```

## Tipos de convenio soportados
| Tipo | Remunerado | Quién paga ARL |
|------|-----------|----------------|
| 01   | Sí        | Organización   |
| 02   | Sí        | EAFIT          |
| 03   | No        | Organización   |
| 04   | No        | EAFIT          |

## Equipo
Reto 2 · Motor de Procesos · Beca IA EAFIT · 2025
