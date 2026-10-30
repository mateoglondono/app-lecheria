"""Modelos SQLAlchemy (ORM) de la base de datos en Supabase."""
from datetime import date, datetime

from sqlalchemy import Date, Float, ForeignKey, Integer, TIMESTAMP, UniqueConstraint, func
from sqlalchemy.dialects.postgresql import JSONB
from sqlalchemy.orm import DeclarativeBase, Mapped, mapped_column


class Base(DeclarativeBase):
    pass


class Periodo(Base):
    # Tabla existente (creada por migrar.py): semanas de recolección
    __tablename__ = 'periodos'

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    fecha_inicio: Mapped[date] = mapped_column(Date, nullable=False)
    fecha_fin: Mapped[date] = mapped_column(Date, nullable=False)
    precio_aplicado: Mapped[float] = mapped_column(Float, nullable=False)
    creado_en: Mapped[datetime] = mapped_column(TIMESTAMP, nullable=False, server_default=func.now())


class Registro(Base):
    # Tabla existente (creada por migrar.py): ordeños diarios
    __tablename__ = 'registros'
    __table_args__ = (UniqueConstraint('periodo_id', 'fecha'),)

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    periodo_id: Mapped[int] = mapped_column(Integer, ForeignKey('periodos.id', ondelete='CASCADE'), nullable=False)
    fecha: Mapped[date] = mapped_column(Date, nullable=False)
    litros_manana: Mapped[float] = mapped_column(Float, nullable=False, default=0)
    litros_tarde: Mapped[float] = mapped_column(Float, nullable=False, default=0)
    actualizado_en: Mapped[datetime] = mapped_column(TIMESTAMP, nullable=False, server_default=func.now())


class GastoSemanal(Base):
    # Liquidación financiera de un periodo: ingresos, costos, punto de equilibrio y utilidad
    __tablename__ = 'gastos_semanales'

    id: Mapped[int] = mapped_column(Integer, primary_key=True) # SERIAL en PostgreSQL
    # Una sola liquidación por periodo; si se borra el periodo, se borra su liquidación
    periodo_id: Mapped[int] = mapped_column(Integer, ForeignKey('periodos.id', ondelete='CASCADE'),
                                            nullable=False, unique=True)
    litros_totales: Mapped[float] = mapped_column(Float, nullable=False)
    precio_venta_litro: Mapped[float] = mapped_column(Float, nullable=False)
    ingresos_leche: Mapped[float] = mapped_column(Float, nullable=False)
    costo_cuidos: Mapped[float] = mapped_column(Float, nullable=False, default=0)
    # Lista de compras: [{"insumo", "detalle", "cantidad_bultos", "precio_bulto", "subtotal"}, ...]
    detalle_cuidos: Mapped[list] = mapped_column(JSONB, nullable=False, default=list)
    costos_fijos: Mapped[float] = mapped_column(Float, nullable=False, default=0)
    costo_total_semana: Mapped[float] = mapped_column(Float, nullable=False)
    punto_equilibrio_litro: Mapped[float] = mapped_column(Float, nullable=False)
    utilidad_neta_semana: Mapped[float] = mapped_column(Float, nullable=False)
    actualizado_en: Mapped[datetime] = mapped_column(TIMESTAMP, nullable=False, server_default=func.now(),
                                                     onupdate=func.now())
