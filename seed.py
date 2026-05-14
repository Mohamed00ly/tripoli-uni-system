"""Run once after `docker compose up -d` to populate the database."""
from app import app
from models import (db, User, PreRegisteredStudent, Department, Course,
                    GradeFormula, CourseAssignment)
from werkzeug.security import generate_password_hash


DEPARTMENTS = [
    ('القانون',           'Law'),
    ('تقنية المعلومات',   'Information Technology'),
    ('اللغات',            'Languages'),
    ('الهندسة',           'Engineering'),
    ('طب الأسنان',        'Dentistry'),
    ('الصيدلة',           'Pharmacy'),
    ('التصميم الداخلي',   'Interior Design'),
    ('إدارة الأعمال',     'Business Administration'),
]


def seed():
    with app.app_context():
        db.create_all()

        # ── Departments ──────────────────────────────────────────────────────
        if not Department.query.first():
            for name_ar, name_en in DEPARTMENTS:
                db.session.add(Department(name_ar=name_ar, name_en=name_en))
            db.session.commit()
            print('✓ 8 departments added')

        it  = Department.query.filter_by(name_en='Information Technology').first()
        eng = Department.query.filter_by(name_en='Engineering').first()

        # ── Courses with prerequisites ───────────────────────────────────────
        if not Course.query.first():
            math1 = Course(code='IT101', name_ar='رياضيات 1',      name_en='Mathematics 1',    department_id=it.id,  credits=3)
            prog1 = Course(code='IT102', name_ar='برمجة 1',        name_en='Programming 1',    department_id=it.id,  credits=3)
            mech1 = Course(code='ENG101', name_ar='ميكانيكا 1',    name_en='Mechanics 1',      department_id=eng.id, credits=3)
            db.session.add_all([math1, prog1, mech1])
            db.session.commit()

            math2 = Course(code='IT201', name_ar='رياضيات 2',      name_en='Mathematics 2',    department_id=it.id,  prerequisite_id=math1.id, credits=3)
            prog2 = Course(code='IT202', name_ar='برمجة 2',        name_en='Programming 2',    department_id=it.id,  prerequisite_id=prog1.id, credits=3)
            db.session.add_all([math2, prog2])
            db.session.commit()

            db_c = Course(code='IT301', name_ar='قواعد البيانات',  name_en='Databases',        department_id=it.id,  prerequisite_id=prog2.id, credits=3)
            net  = Course(code='IT302', name_ar='شبكات الحاسوب',   name_en='Computer Networks',department_id=it.id,  prerequisite_id=math2.id, credits=3)
            db.session.add_all([db_c, net])
            db.session.commit()
            print('✓ Courses added (with prerequisites)')

        # ── Grade formulas (40% classwork + 60% final) ───────────────────────
        if not GradeFormula.query.first():
            for c in Course.query.all():
                db.session.add(GradeFormula(
                    course_id=c.id, classwork_weight=40, final_weight=60))
            db.session.commit()
            print('✓ Grade formulas: 40% أعمال + 60% امتحان')

        # ── Admin ────────────────────────────────────────────────────────────
        if not User.query.filter_by(role='admin').first():
            admin = User(
                student_id='ADMIN001', full_name='مدير النظام',
                email='admin@tripoli-uni.edu',
                password_hash=generate_password_hash('Admin@2024'),
                role='admin', status='approved', financial_status='paid',
                admin_permissions='full_access',
            )
            db.session.add(admin)
            db.session.commit()
            print('✓ Admin  →  admin@tripoli-uni.edu  /  Admin@2024')

        # ── Lecturer ─────────────────────────────────────────────────────────
        if not User.query.filter_by(role='lecturer').first():
            lec = User(
                student_id='LEC001', full_name='د. سعيد الأمين',
                email='lecturer@tripoli-uni.edu',
                password_hash=generate_password_hash('Lecturer@2024'),
                role='lecturer', status='approved', financial_status='paid',
                department_id=it.id if it else None,
            )
            db.session.add(lec)
            db.session.commit()

            it101 = Course.query.filter_by(code='IT101').first()
            it102 = Course.query.filter_by(code='IT102').first()
            for c in [it101, it102]:
                if c:
                    db.session.add(CourseAssignment(
                        lecturer_id=lec.id, course_id=c.id,
                        semester='1', year=2024))
            db.session.commit()
            print('✓ Lecturer  →  lecturer@tripoli-uni.edu  /  Lecturer@2024')

        # ── Pre-registered students ───────────────────────────────────────────
        if not PreRegisteredStudent.query.first():
            it_dept = Department.query.filter_by(name_en='Information Technology').first()
            pre = [
                PreRegisteredStudent(student_id='218001', full_name='أحمد علي محمد',       department_id=it_dept.id if it_dept else None),
                PreRegisteredStudent(student_id='218002', full_name='فاطمة عمر عبدالله',   department_id=it_dept.id if it_dept else None),
                PreRegisteredStudent(student_id='218003', full_name='محمد حسن علي',        department_id=it_dept.id if it_dept else None),
                PreRegisteredStudent(student_id='218004', full_name='عائشة إبراهيم خالد',  department_id=it_dept.id if it_dept else None),
                PreRegisteredStudent(student_id='218005', full_name='عمر محمود سالم',      department_id=it_dept.id if it_dept else None),
            ]
            db.session.add_all(pre)
            db.session.commit()
            print('✓ 5 pre-registered students (218001 → 218005)')

        print('\n✅ Database seeded!')
        print('   Admin:    admin@tripoli-uni.edu / Admin@2024')
        print('   Lecturer: lecturer@tripoli-uni.edu / Lecturer@2024')


if __name__ == '__main__':
    seed()
