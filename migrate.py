"""Run once to apply all schema changes (Phase 4 + Phase 5)."""
from app import app
from models import db


def migrate():
    with app.app_context():
        # Create any new tables that don't exist yet (e.g. system_settings)
        db.create_all()
        print('✓ db.create_all() — new tables created if missing')

        # Add new columns to existing tables
        with db.engine.connect() as conn:
            for sql in [
                "ALTER TABLE users ADD COLUMN IF NOT EXISTS admin_permissions VARCHAR(50)",
                "ALTER TABLE enrollments ADD COLUMN IF NOT EXISTS dropped_by_admin BOOLEAN DEFAULT FALSE",
            ]:
                try:
                    conn.execute(db.text(sql))
                    conn.commit()
                    print(f'✓ {sql}')
                except Exception as exc:
                    print(f'  skipped ({exc})')

        print('\n✅ Migration complete.')


if __name__ == '__main__':
    migrate()
