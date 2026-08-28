# Despliegue del Sistema

## 1. Supabase
1. Crea un proyecto en Supabase.
2. Abre la configuración de conexión de PostgreSQL.
3. Copia la cadena de conexión URI.
4. Guárdala para usarla como `DATABASE_URL`.

No es necesario crear manualmente las tablas: `app.py` las crea al iniciar.

## 2. GitHub
1. Crea un repositorio nuevo.
2. Sube estos archivos:
   - app.py
   - requirements.txt
   - render.yaml
   - .gitignore
3. No subas tu contraseña ni la URL de la base de datos.

## 3. Render
1. Crea un nuevo Web Service desde tu repositorio de GitHub.
2. Render detectará Python.
3. Configura las variables de entorno:
   - DATABASE_URL = URL PostgreSQL de Supabase
   - ADMIN_USER = nombre del administrador inicial
   - ADMIN_PASSWORD = contraseña inicial segura
4. Build Command:
   pip install -r requirements.txt
5. Start Command:
   streamlit run app.py --server.address 0.0.0.0 --server.port $PORT

Al iniciar por primera vez, la aplicación crea sus tablas y el usuario administrador inicial.

## Importante
El código ya no usa SQLite ni el archivo docentes.db. Los datos quedan almacenados en PostgreSQL de Supabase.
