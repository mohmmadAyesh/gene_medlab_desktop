from database import engine,Base
import models

models.Base.metadata.create_all(bind=engine)
print("All tables created or already exist")