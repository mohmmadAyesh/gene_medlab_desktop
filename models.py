from sqlalchemy import (
    Column, Integer, String, Date, DateTime, Enum, ForeignKey, Table
)
from sqlalchemy.orm import relationship, declarative_base

from database import Base

class Role(Base):
    __tablename__ = 'roles'
    role_id = Column(Integer, primary_key=True)
    name    = Column(String, unique=True, nullable=False)
    users   = relationship("User", back_populates="role")

class User(Base):
    __tablename__ = 'users'
    user_id       = Column(Integer, primary_key=True)
    username      = Column(String, unique=True, nullable=False)
    password_hash = Column(String, nullable=False)
    role_id       = Column(Integer, ForeignKey('roles.role_id'))
    role          = relationship("Role", back_populates="users")
    patient       = relationship("PatientProfile", uselist=False, back_populates="user")

class PatientProfile(Base):
    __tablename__ = 'patient_profiles'
    patient_id = Column(Integer, ForeignKey('users.user_id'), primary_key=True)
    first_name = Column(String)
    last_name  = Column(String)
    dob        = Column(Date)
    user       = relationship("User", back_populates="patient")
    conditions = relationship("PatientCondition", back_populates="patient")
    preferences= relationship("Preference", back_populates="patient")
    exclusions = relationship("UserExclusion", back_populates="patient")
    plans      = relationship("MealPlan", back_populates="patient")

class HealthCondition(Base):
    __tablename__ = 'health_conditions'
    condition_id = Column(Integer, primary_key=True)
    name         = Column(String, unique=True, nullable=False)
    patient_links= relationship("PatientCondition", back_populates="condition")
    rules        = relationship("ExclusionRule", back_populates="condition")

class PatientCondition(Base):
    __tablename__ = 'patient_conditions'
    patient_id   = Column(Integer, ForeignKey('patient_profiles.patient_id'), primary_key=True)
    condition_id = Column(Integer, ForeignKey('health_conditions.condition_id'), primary_key=True)
    patient      = relationship("PatientProfile", back_populates="conditions")
    condition    = relationship("HealthCondition", back_populates="patient_links")

class MealItem(Base):
    __tablename__ = 'meal_items'
    item_id   = Column(Integer, primary_key=True)
    name      = Column(String, nullable=False)
    eat_time  = Column(Enum('Breakfast','Lunch','Dinner', name='eat_time'))
    group     = Column(Integer, nullable=False)
    color     = Column(String, nullable=True)
    rules     = relationship("ExclusionRule", back_populates="item")
    user_ex   = relationship("UserExclusion", back_populates="item")
    plan_ent  = relationship("MealPlanEntry", back_populates="item")

class ExclusionRule(Base):
    __tablename__ = 'exclusion_rules'
    rule_id      = Column(Integer, primary_key=True)
    condition_id = Column(Integer, ForeignKey('health_conditions.condition_id'))
    item_id      = Column(Integer, ForeignKey('meal_items.item_id'))
    condition    = relationship("HealthCondition", back_populates="rules")
    item         = relationship("MealItem", back_populates="rules")
    overrides    = relationship("UserExclusion", back_populates="rule")

class UserExclusion(Base):
    __tablename__ = 'user_exclusions'
    exclusion_id = Column(Integer, primary_key=True)
    patient_id   = Column(Integer, ForeignKey('patient_profiles.patient_id'))
    item_id      = Column(Integer, ForeignKey('meal_items.item_id'))
    rule_id      = Column(Integer, ForeignKey('exclusion_rules.rule_id'), nullable=True)
    patient      = relationship("PatientProfile", back_populates="exclusions")
    item         = relationship("MealItem", back_populates="user_ex")
    rule         = relationship("ExclusionRule", back_populates="overrides")

class Preference(Base):
    __tablename__ = 'preferences'
    pref_id    = Column(Integer, primary_key=True)
    patient_id = Column(Integer, ForeignKey('patient_profiles.patient_id'))
    # add your extra pref columns here, e.g. calorie_target, allergies, etc.
    patient    = relationship("PatientProfile", back_populates="preferences")
    meal_name = Column(String,nullable=False)
    rating = Column(String,nullable=True)
class MealPlan(Base):
    __tablename__ = 'meal_plans'
    plan_id      = Column(Integer, primary_key=True)
    patient_id   = Column(Integer, ForeignKey('patient_profiles.patient_id'))
    generated_at = Column(DateTime)
    entries      = relationship("MealPlanEntry", back_populates="plan")
    patient      = relationship("PatientProfile", back_populates="plans")

class MealPlanEntry(Base):
    __tablename__ = 'meal_plan_entries'
    entry_id   = Column(Integer, primary_key=True)
    plan_id    = Column(Integer, ForeignKey('meal_plans.plan_id'))
    day_number = Column(Integer)
    eat_time   = Column(Enum('Breakfast','Lunch','Dinner', name='eat_time'))
    item_id    = Column(Integer, ForeignKey('meal_items.item_id'))
    plan       = relationship("MealPlan", back_populates="entries")
    item       = relationship("MealItem", back_populates="plan_ent")
