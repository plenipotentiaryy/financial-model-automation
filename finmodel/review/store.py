"""SQLite revisions and append-only audit. Provider calls do not hold transactions."""
from contextlib import contextmanager
from dataclasses import asdict
import json
from pathlib import Path
import sqlite3
from datetime import datetime, timezone

from ..schemas import ModelError
from .contracts import Candidate


def now():
    return datetime.now(timezone.utc).isoformat()


class Store:
    def __init__(self, run, create=False):
        self.root = Path(run).resolve()
        self.path = self.root / '.finmodel' / 'review.sqlite3'
        if not create and not self.path.is_file():
            raise ModelError(f'Review run does not exist: {self.root}')
        self.path.parent.mkdir(parents=True, exist_ok=True)
        with self.connect() as db:
            db.executescript('''
            CREATE TABLE IF NOT EXISTS candidates(id TEXT, revision INTEGER, payload TEXT NOT NULL,
              state TEXT NOT NULL, reason TEXT NOT NULL, PRIMARY KEY(id, revision));
            CREATE TABLE IF NOT EXISTS events(seq INTEGER PRIMARY KEY, at TEXT, kind TEXT, payload TEXT);
            CREATE TABLE IF NOT EXISTS cache(key TEXT PRIMARY KEY, payload TEXT);
            CREATE TABLE IF NOT EXISTS requests(id INTEGER PRIMARY KEY, at TEXT, reserved TEXT, actual TEXT, payload TEXT);
            ''')

    @contextmanager
    def connect(self):
        with sqlite3.connect(self.path, timeout=20) as db:
            db.row_factory = sqlite3.Row
            yield db

    def event(self, db, kind, payload):
        db.execute('INSERT INTO events(at,kind,payload) VALUES(?,?,?)', (now(), kind, json.dumps(payload, ensure_ascii=False)))

    def add(self, candidate):
        with self.connect() as db:
            db.execute('BEGIN IMMEDIATE')
            if db.execute('SELECT 1 FROM candidates WHERE id=?', (candidate.candidate_id,)).fetchone():
                return False
            db.execute('INSERT INTO candidates VALUES(?,?,?,?,?)', (candidate.candidate_id, 1, json.dumps(candidate.payload()), 'pending', 'Not verified'))
            self.event(db, 'extracted', candidate.payload())
            return True

    def list(self):
        with self.connect() as db:
            rows = db.execute('SELECT c.* FROM candidates c JOIN (SELECT id, MAX(revision) revision FROM candidates GROUP BY id) latest USING(id,revision) ORDER BY c.id').fetchall()
        return [dict(candidate=Candidate(**json.loads(r['payload'])), state=r['state'], reason=r['reason']) for r in rows]

    def get(self, cid):
        for row in self.list():
            if row['candidate'].candidate_id == cid:
                return row
        raise ModelError('Unknown candidate')

    def set_state(self, cid, revision, state, reason, event=None, db=None, expected_states=None):
        if db is None:
            with self.connect() as conn:
                conn.execute('BEGIN IMMEDIATE')
                return self.set_state(cid, revision, state, reason, event, conn, expected_states)
        latest = db.execute('SELECT MAX(revision) FROM candidates WHERE id=?', (cid,)).fetchone()[0]
        if latest != revision:
            raise ModelError('Stale revision; reload the candidate')
        current = db.execute('SELECT state FROM candidates WHERE id=? AND revision=?', (cid, revision)).fetchone()[0]
        if expected_states is not None and current not in expected_states:
            return False
        db.execute('UPDATE candidates SET state=?, reason=? WHERE id=? AND revision=?', (state, reason, cid, revision))
        self.event(db, 'state', {'candidate_id': cid, 'revision': revision, 'state': state, 'reason': reason, 'detail': event})

    def revise(self, cid, revision, changes, reason):
        allowed = {'metric', 'company_id', 'company_name', 'fiscal_year', 'currency', 'monetary_unit', 'scope', 'period_type'}
        if not changes or set(changes) - allowed:
            raise ModelError('Only mapping/context may be corrected; re-extract changed source values')
        candidate = self.get(cid)['candidate']
        if candidate.revision != revision:
            raise ModelError('Stale revision')
        if any((type(v) is not int if k == 'fiscal_year' else not isinstance(v, str)) for k, v in changes.items()):
            raise ModelError('Mapping values must be text; fiscal_year must be an integer')
        data = candidate.payload()
        data.update(changes)
        data.update(revision=revision + 1, normalized_value=None)
        new = Candidate(**data)
        from .gate import inspect_candidate
        inspect_candidate(new)
        with self.connect() as db:
            db.execute('BEGIN IMMEDIATE')
            if db.execute('SELECT MAX(revision) FROM candidates WHERE id=?', (cid,)).fetchone()[0] != revision:
                raise ModelError('Stale revision')
            db.execute('INSERT INTO candidates VALUES(?,?,?,?,?)', (cid, new.revision, json.dumps(new.payload()), 'pending', 'Mapping revised; prior verification invalid'))
            self.event(db, 'revision', {'before': candidate.payload(), 'after': new.payload(), 'reason': reason})
        return new

    def history(self):
        with self.connect() as db:
            return [dict(r) | {'payload': json.loads(r['payload'])} for r in db.execute('SELECT * FROM events ORDER BY seq')]

    def cached(self, key):
        with self.connect() as db:
            row = db.execute('SELECT payload FROM cache WHERE key=?', (key,)).fetchone()
            return json.loads(row[0]) if row else None

    def cache(self, key, result):
        with self.connect() as db:
            db.execute('INSERT OR REPLACE INTO cache VALUES(?,?)', (key, json.dumps(result)))
