from flask_sqlalchemy import SQLAlchemy
from flask_login import UserMixin
from datetime import datetime

db = SQLAlchemy()


class PreRegisteredStudent(db.Model):
    __tablename__ = 'pre_registered_students'
    id = db.Column(db.Integer, primary_key=True)
    student_id = db.Column(db.String(50), unique=True, nullable=False)
    full_name = db.Column(db.String(200), nullable=False)
    birth_date = db.Column(db.Date, nullable=True)
    department_id = db.Column(db.Integer, db.ForeignKey('departments.id'), nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    department = db.relationship('Department', backref='pre_registered')
    documents = db.relationship('Document', backref='pre_reg_student', lazy='dynamic',
                                foreign_keys='Document.pre_reg_id')


class Department(db.Model):
    __tablename__ = 'departments'
    id = db.Column(db.Integer, primary_key=True)
    name_ar = db.Column(db.String(200), nullable=False)
    name_en = db.Column(db.String(200), nullable=False)
    courses = db.relationship('Course', backref='department', lazy='dynamic')


class User(UserMixin, db.Model):
    __tablename__ = 'users'
    id = db.Column(db.Integer, primary_key=True)
    student_id = db.Column(db.String(20), unique=True, nullable=False)
    full_name = db.Column(db.String(200), nullable=False)
    email = db.Column(db.String(200), unique=True, nullable=False)
    password_hash = db.Column(db.String(256), nullable=False)
    role = db.Column(db.String(20), default='student')   # admin | lecturer | student
    status = db.Column(db.String(20), default='pending') # pending | approved | rejected
    financial_status = db.Column(db.String(20), default='unpaid')
    department_id = db.Column(db.Integer, db.ForeignKey('departments.id'), nullable=True)
    birth_date = db.Column(db.Date, nullable=True)
    profile_image = db.Column(db.String(300), nullable=True)
    admin_permissions = db.Column(db.String(50), nullable=True)  # full_access | manage_users | manage_grades
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    department = db.relationship('Department', backref='students')
    enrollments = db.relationship('Enrollment', backref='student', lazy='dynamic',
                                  foreign_keys='Enrollment.user_id')
    assignments = db.relationship('CourseAssignment', backref='lecturer', lazy='dynamic',
                                  foreign_keys='CourseAssignment.lecturer_id')
    documents = db.relationship('Document', backref='user', lazy='dynamic',
                                foreign_keys='Document.user_id')

    @property
    def is_admin(self):
        return self.role == 'admin'

    @property
    def is_lecturer(self):
        return self.role == 'lecturer'

    @property
    def is_student(self):
        return self.role == 'student'

    @property
    def has_full_access(self):
        # NULL permissions = full access (backward-compat for primary admin)
        return self.role == 'admin' and (
            self.admin_permissions is None or self.admin_permissions == 'full_access'
        )


class Course(db.Model):
    __tablename__ = 'courses'
    id = db.Column(db.Integer, primary_key=True)
    code = db.Column(db.String(20), unique=True, nullable=False)
    name_ar = db.Column(db.String(200), nullable=False)
    name_en = db.Column(db.String(200), nullable=False)
    department_id = db.Column(db.Integer, db.ForeignKey('departments.id'), nullable=False)
    prerequisite_id = db.Column(db.Integer, db.ForeignKey('courses.id'), nullable=True)
    credits = db.Column(db.Integer, default=3)

    prerequisite = db.relationship('Course', remote_side=[id], backref='next_courses')
    enrollments = db.relationship('Enrollment', backref='course', lazy='dynamic')
    grade_formula = db.relationship('GradeFormula', backref='course', uselist=False)
    schedules = db.relationship('ClassSchedule', backref='course', lazy='dynamic')
    assignments = db.relationship('CourseAssignment', backref='course', lazy='dynamic')


class GradeFormula(db.Model):
    __tablename__ = 'grade_formulas'
    id = db.Column(db.Integer, primary_key=True)
    course_id = db.Column(db.Integer, db.ForeignKey('courses.id'), unique=True, nullable=False)
    classwork_weight = db.Column(db.Integer, default=40)
    final_weight = db.Column(db.Integer, default=60)


class ClassSchedule(db.Model):
    __tablename__ = 'class_schedules'
    id = db.Column(db.Integer, primary_key=True)
    course_id = db.Column(db.Integer, db.ForeignKey('courses.id'), nullable=False)
    day_of_week = db.Column(db.String(20), nullable=False)
    start_time = db.Column(db.String(10), nullable=False)
    end_time = db.Column(db.String(10), nullable=False)
    room = db.Column(db.String(50))
    semester = db.Column(db.String(10))
    year = db.Column(db.Integer)


class CourseAssignment(db.Model):
    __tablename__ = 'course_assignments'
    id = db.Column(db.Integer, primary_key=True)
    lecturer_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=False)
    course_id = db.Column(db.Integer, db.ForeignKey('courses.id'), nullable=False)
    semester = db.Column(db.String(10), nullable=False)
    year = db.Column(db.Integer, nullable=False)
    __table_args__ = (
        db.UniqueConstraint('lecturer_id', 'course_id', 'semester', 'year', name='uq_assignment'),
    )


class Enrollment(db.Model):
    __tablename__ = 'enrollments'
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=False)
    course_id = db.Column(db.Integer, db.ForeignKey('courses.id'), nullable=False)
    semester = db.Column(db.String(10), nullable=False)
    year = db.Column(db.Integer, nullable=False)
    classwork_grade = db.Column(db.Float, nullable=True)
    final_grade = db.Column(db.Float, nullable=True)
    grade = db.Column(db.Float, nullable=True)   # computed total
    grade_locked = db.Column(db.Boolean, default=False)
    dropped_by_admin = db.Column(db.Boolean, default=False)
    status = db.Column(db.String(20), default='enrolled')
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    __table_args__ = (
        db.UniqueConstraint('user_id', 'course_id', 'semester', 'year', name='uq_enrollment'),
    )


class Document(db.Model):
    __tablename__ = 'documents'
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=True)
    pre_reg_id = db.Column(db.Integer, db.ForeignKey('pre_registered_students.id'), nullable=True)
    document_type = db.Column(db.String(50), nullable=False)  # birth_certificate | qualification
    filename = db.Column(db.String(300), nullable=False)
    original_name = db.Column(db.String(300), nullable=False)
    uploaded_at = db.Column(db.DateTime, default=datetime.utcnow)


class SystemSettings(db.Model):
    __tablename__ = 'system_settings'
    id         = db.Column(db.Integer, primary_key=True)
    logo_path  = db.Column(db.String(300), nullable=True)   # relative to static/
    site_name  = db.Column(db.String(200), default='جامعة طرابلس')
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)


class AuditLog(db.Model):
    __tablename__ = 'audit_logs'
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=True)
    action = db.Column(db.String(200), nullable=False)
    entity_type = db.Column(db.String(50))
    entity_id = db.Column(db.Integer)
    old_value = db.Column(db.Text)
    new_value = db.Column(db.Text)
    ip_address = db.Column(db.String(50))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    actor = db.relationship('User', foreign_keys=[user_id])
