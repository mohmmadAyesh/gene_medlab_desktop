# database.py
import os
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker, scoped_session, declarative_base

DATABASE_URL = os.getenv(
    "DATABASE_URL",
    "postgresql://mealplanner_user:ayesh123@127.0.0.1:5432/mealplanner_db"
)

# 1) The Engine (talks to Postgres)
engine = create_engine(DATABASE_URL, echo=True, future=True)

# 2) Session factory
SessionLocal = sessionmaker(
    autocommit=False,
    autoflush=False,
    bind=engine,
    future=True
)

# 3) Optional thread-local scoped session
db_session = scoped_session(SessionLocal)

# 4) Declarative base for models
Base = declarative_base()
