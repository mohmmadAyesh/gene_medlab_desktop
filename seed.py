# seeds.py
from database import engine,SessionLocal,Base
from models import MealItem,HealthCondition,ExclusionRule,Role,User,PatientProfile,PatientCondition
from sqlalchemy.exc import IntegrityError
import bcrypt,datetime
def seed_meal_items():
    Base.metadata.create_all(bind=engine)
    session = SessionLocal()

    # List of items to seed
    items = [
            {"name": "١-٢ بيضة مسلوقة", "color": "A", "eat_time": "Breakfast", "group": 1},
    {"name": "١-٢ علبة لبن رايب", "color": "A", "eat_time": "Breakfast", "group": 1},
    {"name": "٥ ملاعق جبنة فيتا + خيار", "color": "A", "eat_time": "Breakfast", "group": 1},
    {"name": "٤ قطع جبنة مثلثات", "color": "C", "eat_time": "Breakfast", "group": 1},
    {"name": "قطعة جبنة صفراء كيه٢", "color": "B", "eat_time": "Breakfast", "group": 1},
    {"name": "جبنة قريش + بندورة + خيار", "color": "A", "eat_time": "Breakfast", "group": 1},
    {"name": "١-٢ حبة أفوكادو بدون خبز", "color": "A", "eat_time": "Breakfast", "group": 1},
    {"name": "١ صحن سلطة بدون خبز", "color": "A", "eat_time": "Breakfast", "group": 1},
    {"name": "٥٠ غرام مكسرات (فستق-لوز-عين جمل)", "color": "B", "eat_time": "Breakfast", "group": 1},
    {"name": "٤ ملاعق لبنة كبيرة + ٢ حبة خيار بدون خبز", "color": "A", "eat_time": "Breakfast", "group": 1},
    {"name": "٢ بيضة مقلية بزيت الزيتون + خضار بدون خبز", "color": "A", "eat_time": "Breakfast", "group": 1},
    {"name": "١ صحن متبل باتنجان", "color": "C", "eat_time": "Breakfast", "group": 1},
    {"name": "٣ ملاعق كبيرة حمص + خضار أو سلطة", "color": "C", "eat_time": "Breakfast", "group": 1},
    {"name": "١ تفاحة", "color": "A", "eat_time": "Breakfast", "group": 2},
    {"name": "خبز شعير ٥٠ جرام بيتا أو فرشوحة", "color": "B", "eat_time": "Breakfast", "group": 2},
    {"name": "٢ حبة خوخ", "color": "A", "eat_time": "Breakfast", "group": 2},
    {"name": "٢ حبة بندورة", "color": "A", "eat_time": "Breakfast", "group": 2},
    {"name": "٢ حبة خيار", "color": "A", "eat_time": "Breakfast", "group": 2},
    {"name": "٢ شريحة بطيخ", "color": "A", "eat_time": "Breakfast", "group": 2},
    {"name": "٢ شريحة شمام", "color": "A", "eat_time": "Breakfast", "group": 2},
    {"name": "٢ حبة أجاص", "color": "A", "eat_time": "Breakfast", "group": 2},
    {"name": "حبة جوافة", "color": "A", "eat_time": "Breakfast", "group": 2},
    {"name": "٦ حبات فراولة", "color": "A", "eat_time": "Breakfast", "group": 2},
    {"name": "جريب فروت", "color": "A", "eat_time": "Breakfast", "group": 2},
    {"name": "برتقالة", "color": "A", "eat_time": "Breakfast", "group": 2},
    {"name": "نصف حبة بوملي", "color": "A", "eat_time": "Breakfast", "group": 2},
    {"name": "حبة فرمسون", "color": "A", "eat_time": "Breakfast", "group": 2},
    {"name": "٢ حبة كيوي", "color": "A", "eat_time": "Breakfast", "group": 2},
    {"name": "٢٥٠ غرام صدر دجاج مشوي", "color": "A", "eat_time": "Lunch", "group": 1},
    {"name": "٢٥٠ غرام سمك أو تونة", "color": "A", "eat_time": "Lunch", "group": 1},
    {"name": "٢٥٠ غرام لحم عجل مشوي بدون دهون", "color": "A", "eat_time": "Lunch", "group": 1},
    {"name": "١ علبة تونة مصفاة من الزيت", "color": "A", "eat_time": "Lunch", "group": 1},
    {"name": "١ صحن شوربة عدس", "color": "B", "eat_time": "Lunch", "group": 1},
    {"name": "١ صحن فول مدمس", "color": "A", "eat_time": "Lunch", "group": 1},
    {"name": "١ صحن طبيخ بازيلاء بدون أرز", "color": "A", "eat_time": "Lunch", "group": 1},
    {"name": "١ صحن طبيخ فاصولياء خضراء بدون أرز", "color": "A", "eat_time": "Lunch", "group": 1},
    {"name": "١ صحن سبانخ بدون أرز", "color": "A", "eat_time": "Lunch", "group": 1},
    {"name": "١ صحن طبيخ كوسا مقطعة بدون أرز", "color": "A", "eat_time": "Lunch", "group": 1},
    {"name": "١ صحن طبيخ فول أخضر بدون أرز", "color": "A", "eat_time": "Lunch", "group": 1},
    {"name": "١ صحن طبيخ ملوخية بدون أرز", "color": "A", "eat_time": "Lunch", "group": 1},
    {"name": "١ صحن طبيخ بامية بدون أرز", "color": "A", "eat_time": "Lunch", "group": 1},
    {"name": "١ صحن طبيخ يقطين بدون أرز", "color": "A", "eat_time": "Lunch", "group": 1},
    {"name": "٢٥٠ غرام كبدة دجاج", "color": "A", "eat_time": "Lunch", "group": 1},
    {"name": "٢٥٠ غرام شاورما بيتية خالي الدهن", "color": "B", "eat_time": "Lunch", "group": 1},
    {"name": "٢٥٠ غرام شيش طاووق", "color": "A", "eat_time": "Lunch", "group": 1},
    {"name": "٢٥٠ غرام كباب خالي الدهن", "color": "A", "eat_time": "Lunch", "group": 1},
    {"name": "٢٥٠ غرام طحال مشوي", "color": "A", "eat_time": "Lunch", "group": 1},
    {"name": "٢٥٠ غرام فطر + لحمة + بصل", "color": "B", "eat_time": "Lunch", "group": 1},
    {"name": "٢٥٠ غرام فطر + بيض + بصل", "color": "B", "eat_time": "Lunch", "group": 1},
    {"name": "١ صحن لوبياء بدون أرز", "color": "B", "eat_time": "Lunch", "group": 1},
    {"name": "١ صحن كوسا مغشي بدون خبز", "color": "A", "eat_time": "Lunch", "group": 1},
    {"name": "١ صحن مسقعة بائنجان بدون خبز", "color": "B", "eat_time": "Lunch", "group": 1},
    {"name": "١ صحن خبيزة بدون أرز", "color": "A", "eat_time": "Lunch", "group": 1},
    {"name": "١ صحن سلق بدون أرز", "color": "A", "eat_time": "Lunch", "group": 1},
    {"name": "١ صحن زهرة بلبن بدون خبز", "color": "A", "eat_time": "Lunch", "group": 1},
    {"name": "١ صحن فول أخضر بدون خبز", "color": "A", "eat_time": "Lunch", "group": 1},
    {"name": "١ صحن زهرة بندورة بدون خبز", "color": "A", "eat_time": "Lunch", "group": 1},
    {"name": "١ صحن متبل باذنجان مع ٥٠ غرام خبز شعير", "color": "C", "eat_time": "Lunch", "group": 1},
    {"name": "١ صحن سماقية بدون خبز", "color": "C", "eat_time": "Lunch", "group": 1},
    {"name": "١ صحن رمانية بدون خبز", "color": "C", "eat_time": "Lunch", "group": 1},
    {"name": "٢٥٠ غرام كفتة سمك", "color": "A", "eat_time": "Lunch", "group": 1},
    {"name": "٢٥٠ غرام سمك مقلي", "color": "A", "eat_time": "Lunch", "group": 1},
    {"name": "٢٥٠ غرام لحم خروف", "color": "C", "eat_time": "Lunch", "group": 1},
    {"name": "٢٥٠ غرام كباب مشوي", "color": "A", "eat_time": "Lunch", "group": 1},
    {"name": "٢٥٠ غرام كفتة بدون خبز", "color": "A", "eat_time": "Lunch", "group": 1},
    {"name": "٢٥٠ غرام تايلندي + خضار بدون خبز", "color": "A", "eat_time": "Lunch", "group": 1},
    {"name": "٢٥٠ غرام شاورما بدون خبز", "color": "A", "eat_time": "Lunch", "group": 1},
    {"name": "٢٥٠ غرام جناحين مشوية", "color": "A", "eat_time": "Lunch", "group": 1},
    {"name": "٢٥٠ غرام فخد دجاج", "color": "A", "eat_time": "Lunch", "group": 1},
    {"name": "٢٥٠ غرام كبدة دجاج أو عجل بدون خبز", "color": "A", "eat_time": "Lunch", "group": 1},
    {"name": "٢٥٠ غرام طحال أو فشة بدون خبز", "color": "A", "eat_time": "Lunch", "group": 1},
    {"name": "٢٥٠ غرام حبش", "color": "A", "eat_time": "Lunch", "group": 1},
    {"name": "٢٥٠ غرام ستيكات دجاج", "color": "A", "eat_time": "Lunch", "group": 1},
    {"name": "علبة فطر مع بيض و بصل", "color": "A", "eat_time": "Lunch", "group": 1},
    {"name": "١ صحن ملوخية ورق بدون خبز", "color": "A", "eat_time": "Lunch", "group": 1},
    {"name": "صحن طبيخ حماصيص", "color": "A", "eat_time": "Lunch", "group": 1},
    {"name": "صحن طبيخ رجلة أو بقلة", "color": "A", "eat_time": "Lunch", "group": 1},
    {"name": "صحن طبيخ سلق و عدس", "color": "A", "eat_time": "Lunch", "group": 1},
    {"name": "شوربة خضار", "color": "A", "eat_time": "Lunch", "group": 1},
    {"name": "سلطة خضراء بدون ذرة", "color": "A", "eat_time": "Lunch", "group": 2},
    {"name": "شوربة خضار بدون بطاطس أو ذرة", "color": "A", "eat_time": "Lunch", "group": 2},
    {"name": "شوربة ملفوف", "color": "A", "eat_time": "Lunch", "group": 2},
    {"name": "شوربة فطر و بصل بدون كريمة", "color": "A", "eat_time": "Lunch", "group": 2},
    {"name": "خضار بدون سلطة أو شوربة", "color": "A", "eat_time": "Lunch", "group": 2},
    {"name": "١ صحن بائنجان مقلي", "color": "A", "eat_time": "Dinner", "group": 1},
    {"name": "١ صحن متبل باذنجان", "color": "A", "eat_time": "Dinner", "group": 1},
    {"name": "١ صحن خيار + لبن", "color": "A", "eat_time": "Dinner", "group": 1},
    {"name": "١ صحن سلطة ملفوف مع مايونيز أو خل أو زيت", "color": "B", "eat_time": "Dinner", "group": 1},
    {"name": "١ صحن سلطة خضراء + طحينة", "color": "B", "eat_time": "Dinner", "group": 1},
    {"name": "١ صحن شكشوكة بدون خبز", "color": "A", "eat_time": "Dinner", "group": 1},
    {"name": "١٠٠ جرام فستق حلبي", "color": "B", "eat_time": "Dinner", "group": 1},
    {"name": "١٠٠ جرام فستق سوداني", "color": "B", "eat_time": "Dinner", "group": 1},
    {"name": "١٠٠ غرام لوز", "color": "B", "eat_time": "Dinner", "group": 1},
    {"name": "١٠٠ غرام بندق", "color": "B", "eat_time": "Dinner", "group": 1},
    {"name": "١٠٠ غرام كاجو", "color": "B", "eat_time": "Dinner", "group": 1},
    {"name": "١٠٠ غرام عين جمل", "color": "B", "eat_time": "Dinner", "group": 1},
    {"name": "١٠٠ غرام بذر عين شمس", "color": "B", "eat_time": "Dinner", "group": 1},
    {"name": "صحن عجة بيض + بقدونس + بصل", "color": "A", "eat_time": "Dinner", "group": 1},
    {"name": "٣ أو ٤ حبات فاكهة صنف واحد (تفاح، مشمس، كيوي، أجاص، خوخ)", "color": "A", "eat_time": "Dinner", "group": 1},
    {"name": "علبة فول مدمس", "color": "A", "eat_time": "Dinner", "group": 1},
    {"name": "٢ شريحة بطيخ", "color": "A", "eat_time": "Dinner", "group": 1},
    {"name": "٢ شريحة شمام", "color": "A", "eat_time": "Dinner", "group": 1},
    {"name": "١٠٠ غرام عنب", "color": "A", "eat_time": "Dinner", "group": 1},
    {"name": "٢ حبة تين", "color": "A", "eat_time": "Dinner", "group": 1},
    {"name": "١٠٠ غرام مانجا", "color": "A", "eat_time": "Dinner", "group": 1},
    {"name": "حبة جوافة", "color": "A", "eat_time": "Dinner", "group": 1},
    {"name": "١ صحن سلطة يونانية (سلطة خضراء + جبنة بيضاء قليلة الدسم)", "color": "A", "eat_time": "Dinner", "group": 1}
        # … include the rest of your items here …
    ]
    new_count = 0;
    # Insert into DB, skipping duplicates by name
    for itm in items:
        if not session.query(MealItem).filter_by(name=itm['name']).first():
            session.add(MealItem(
                name=itm['name'],
                eat_time = itm['eat_time'],
                group = itm['group'],
                color = itm.get('color')
            ))
            new_count += 1
    try:
        session.commit()
    except IntegrityError:
        session.rollback()
    finally:
        session.close()
    print(f"Seeded {new_count} new meal items.")
def seed_conditions_and_rules():
    session = SessionLocal()
    names = ['Healthy','Diabetes','Kidney Disease']
    for name in names:
        if not session.query(HealthCondition).filter_by(name=name).first():
            session.add(HealthCondition(name=name))
    session.commit()
    cond_map = {h.name:h.condition_id
                for h in session.query(HealthCondition).all()}
    # exisitng arrays
    from meal_items import DIABETES_EXCLUDED_FOODS,KIDNEY_EXCLUDED_FOODS
    for arr, cond_name in ((DIABETES_EXCLUDED_FOODS,'Diabetes'),(KIDNEY_EXCLUDED_FOODS,'Kidney Disease')):
        cid = cond_map[cond_name]
        for item in arr:
            mi = session.query(MealItem).filter_by(name=item['name']).first()
            if mi and not session.query(ExclusionRule) \
                .filter_by(condition_id=cid,item_id=mi.item_id) \
                    .first():
                session.add(ExclusionRule(
                    condition_id = cid,
                    item_id = mi.item_id
                ))
    session.commit()
    session.close()
def seed_roles():
    """Create Roles and HealthCondition once"""
    Base.metadata.create_all(bind=engine)
    session = SessionLocal()
    ##1)Role
    for name in ("admin","secretary","patient"):
        if not session.query(Role).filter_by(name=name).first():
            session.add(Role(name=name))
    session.commit()
def seed_users_patient():
    """Create admin, secretary, and 5 patient users + profiles."""
    session = SessionLocal()
    roles = {r.name: r for r in session.query(Role).all()}

    # 1) Admin + Secretary
    for username, pwd, role_name in [
        ('admin',     'adminpass',    'admin'),
        ('secretary', 'secretpass',   'secretary'),
    ]:
        if not session.query(User).filter_by(username=username).first():
            pw_hash = bcrypt.hashpw(pwd.encode(), bcrypt.gensalt()).decode()
            session.add(User(
                username=username,
                password_hash=pw_hash,
                role=roles[role_name]
            ))

    # 2) Five patients
    for i in range(1, 6):
        uname = f"patient{i}"
        if not session.query(User).filter_by(username=uname).first():
            pw_hash = bcrypt.hashpw(f"patient{i}pass".encode(), bcrypt.gensalt()).decode()
            user = User(
                username=uname,
                password_hash=pw_hash,
                role=roles['patient']
            )
            session.add(user)
            session.flush()  # so user.user_id is populated

            # Create their PatientProfile
            dob = datetime.date.today() - datetime.timedelta(days=365*30 + i*100)
            profile = PatientProfile(
                patient_id = user.user_id,
                first_name = f"First{i}",
                last_name  = f"Last{i}",
                dob        = dob
            )
            session.add(profile)

            # Give patient1 a Diabetes condition
            if i == 1:
                cond = session.query(HealthCondition) \
                              .filter_by(name="Diabetes") \
                              .one()
                session.add(PatientCondition(patient=profile, condition=cond))

    # commit everything once at the end
    session.commit()
    session.close()
    print("Seeded admin, secretary, and 5 patient users + profiles.")
if __name__ == '__main__':
    seed_meal_items()
    seed_conditions_and_rules()
    seed_roles()
    seed_users_patient()
