"""Migra lecheria.db (SQLite local) a PostgreSQL en Supabase.

Uso:  .\\venv\\Scripts\\python.exe migrar.py
Lee la URI desde .streamlit/secrets.toml. Se puede ejecutar varias veces sin duplicar datos.
"""
import os
import sqlite3
import tomllib
from contextlib import closing

from sqlalchemy import create_engine, text

BASE = os.path.dirname(os.path.abspath(__file__))
SQLITE_PATH = os.path.join(BASE, 'lecheria.db')

ESQUEMA = """
CREATE TABLE IF NOT EXISTS config (
    id            SERIAL PRIMARY KEY,
    precio_litro  DOUBLE PRECISION NOT NULL
);

CREATE TABLE IF NOT EXISTS periodos (
    id               SERIAL PRIMARY KEY,
    fecha_inicio     DATE NOT NULL,
    fecha_fin        DATE NOT NULL,
    precio_aplicado  DOUBLE PRECISION NOT NULL,
    creado_en        TIMESTAMP NOT NULL DEFAULT now(),
    CHECK (fecha_fin >= fecha_inicio)
);

CREATE TABLE IF NOT EXISTS registros (
    id              SERIAL PRIMARY KEY,
    periodo_id      INTEGER NOT NULL REFERENCES periodos(id) ON DELETE CASCADE,
    fecha           DATE NOT NULL,
    litros_manana   DOUBLE PRECISION NOT NULL DEFAULT 0,
    litros_tarde    DOUBLE PRECISION NOT NULL DEFAULT 0,
    actualizado_en  TIMESTAMP NOT NULL DEFAULT now(),
    UNIQUE (periodo_id, fecha)  -- un solo registro por día y periodo
);

-- Bloquea el acceso por la API pública de Supabase (clave anon); la app entra como dueño y no se ve afectada
ALTER TABLE config    ENABLE ROW LEVEL SECURITY;
ALTER TABLE periodos  ENABLE ROW LEVEL SECURITY;
ALTER TABLE registros ENABLE ROW LEVEL SECURITY;
"""


def leer_uri():
    with open(os.path.join(BASE, '.streamlit', 'secrets.toml'), 'rb') as f:
        uri = tomllib.load(f)['connections']['postgresql']['url']
    if not uri.startswith('postgresql'):
        raise SystemExit("Falta la URI de Supabase en .streamlit/secrets.toml")
    return uri


def leer_sqlite():
    with closing(sqlite3.connect(SQLITE_PATH)) as conn:
        conn.row_factory = sqlite3.Row
        tablas = {t: [dict(r) for r in conn.execute(f"SELECT * FROM {t} ORDER BY id")]
                  for t in ('config', 'periodos', 'registros')}
    return tablas


# Conteos y suma de litros para comparar origen y destino
VERIFICACION = """
    SELECT (SELECT COUNT(*) FROM config), (SELECT COUNT(*) FROM periodos), (SELECT COUNT(*) FROM registros),
           (SELECT ROUND(CAST(COALESCE(SUM(litros_manana + litros_tarde), 0) AS NUMERIC), 2) FROM registros)
"""


def main():
    datos = leer_sqlite()
    print("Origen (SQLite):", {t: len(f) for t, f in datos.items()})

    engine = create_engine(leer_uri(), connect_args={'sslmode': 'require'})
    with engine.begin() as conn: # Todo en una transacción: si algo falla, no queda nada a medias
        conn.execute(text(ESQUEMA))

        conn.execute(text("INSERT INTO config (id, precio_litro) VALUES (:id, :precio_litro) "
                          "ON CONFLICT (id) DO NOTHING"), datos['config'])
        conn.execute(text("INSERT INTO periodos (id, fecha_inicio, fecha_fin, precio_aplicado) "
                          "VALUES (:id, CAST(:fecha_inicio AS DATE), CAST(:fecha_fin AS DATE), :precio_aplicado) "
                          "ON CONFLICT (id) DO NOTHING"), datos['periodos'])
        if datos['registros']:
            conn.execute(text("INSERT INTO registros (id, periodo_id, fecha, litros_manana, litros_tarde) "
                              "VALUES (:id, :periodo_id, CAST(:fecha AS DATE), :litros_manana, :litros_tarde) "
                              "ON CONFLICT DO NOTHING"), datos['registros'])

        # Como insertamos IDs explícitos, el contador SERIAL debe continuar desde el máximo
        for tabla in ('config', 'periodos', 'registros'):
            conn.execute(text(f"SELECT setval(pg_get_serial_sequence('{tabla}', 'id'), "
                              f"COALESCE(MAX(id), 1), MAX(id) IS NOT NULL) FROM {tabla}"))

        destino = tuple(conn.execute(text(VERIFICACION)).one())

    origen_litros = round(sum(r['litros_manana'] + r['litros_tarde'] for r in datos['registros']), 2)
    origen = (len(datos['config']), len(datos['periodos']), len(datos['registros']), origen_litros)
    destino = (destino[0], destino[1], destino[2], float(destino[3]))
    print("Origen  (config, periodos, registros, litros):", origen)
    print("Destino (config, periodos, registros, litros):", destino)
    print("✅ Migración verificada" if origen == destino else "⚠️ Los totales NO coinciden, revisa antes de usar la app")


if __name__ == '__main__':
    main()
