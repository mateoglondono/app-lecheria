"""Modelos SQLAlchemy (ORM) de la base de datos en Supabase."""
from datetime import date, datetime

from sqlalchemy import Date, Float, ForeignKey, Integer, String, TIMESTAMP, UniqueConstraint, func
from sqlalchemy.orm import DeclarativeBase, Mapped, mapped_column


class Base(DeclarativeBase):
    pass


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
    # Liquidación financiera de una semana: ingresos, costos, punto de equilibrio y utilidad
    __tablename__ = 'gastos_semanales'

    id: Mapped[int] = mapped_column(Integer, primary_key=True) # SERIAL en PostgreSQL
    fecha_inicio: Mapped[date] = mapped_column(Date, nullable=False)
    fecha_fin: Mapped[date] = mapped_column(Date, nullable=False)
    litros_totales: Mapped[float] = mapped_column(Float, nullable=False)
    precio_venta_litro: Mapped[float] = mapped_column(Float, nullable=False)
    ingresos_leche: Mapped[float] = mapped_column(Float, nullable=False)
    tipo_cuido: Mapped[str] = mapped_column(String(100), nullable=False)
    cantidad_bultos: Mapped[float] = mapped_column(Float, nullable=False, default=0)
    precio_bulto: Mapped[float] = mapped_column(Float, nullable=False, default=0)
    costos_fijos: Mapped[float] = mapped_column(Float, nullable=False, default=0)
    costo_total_semana: Mapped[float] = mapped_column(Float, nullable=False)
    punto_equilibrio_litro: Mapped[float] = mapped_column(Float, nullable=False)
    utilidad_neta_semana: Mapped[float] = mapped_column(Float, nullable=False)
