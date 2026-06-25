import sqlite3
import json
import datetime
from typing import List, Dict

DB_FILE = "pharm_db.sqlite"

def init_db():
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    # 1. 题库表
    cur.execute('''
    CREATE TABLE IF NOT EXISTS question_bank (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        title TEXT,
        full_content TEXT,
        scene TEXT,
        organ_systems TEXT,
        topics TEXT,
        diff_levels TEXT,
        q_types TEXT,
        master_levels TEXT,
        create_time DATETIME
    )
    ''')
    # 2. 药物清单表
    cur.execute('''
    CREATE TABLE IF NOT EXISTS drug_list (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        drug_name TEXT UNIQUE,
        organ_system TEXT,
        drug_class TEXT
    )
    ''')
    # 3. Prompt模板表
    cur.execute('''
    CREATE TABLE IF NOT EXISTS prompt_template (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        template_name TEXT UNIQUE,
        prompt_content TEXT,
        model_type TEXT,
        create_time DATETIME
    )
    ''')
    conn.commit()
    conn.close()

# ---------------- 题库操作 ----------------
def add_question_to_db(data: dict):
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    cur.execute('''
    INSERT INTO question_bank
    (title, full_content, scene, organ_systems, topics, diff_levels, q_types, master_levels, create_time)
    VALUES (?,?,?,?,?,?,?,?,?)
    ''', (
        data["title"],
        data["full_content"],
        data["scene"],
        json.dumps(data["organ_systems"], ensure_ascii=False),
        json.dumps(data["topics"], ensure_ascii=False),
        json.dumps(data["diff_levels"], ensure_ascii=False),
        json.dumps(data["q_types"], ensure_ascii=False),
        json.dumps(data["master_levels"], ensure_ascii=False),
        now
    ))
    conn.commit()
    conn.close()

def query_questions(
        scene=None, organ_system=None, topic=None, diff=None, q_type=None
) -> List[Dict]:
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    sql = "SELECT * FROM question_bank WHERE 1=1"
    params = []
    if scene:
        sql += " AND scene=?"
        params.append(scene)
    if organ_system:
        sql += " AND organ_systems LIKE ?"
        params.append(f'%{organ_system}%')
    if topic:
        sql += " AND topics LIKE ?"
        params.append(f'%{topic}%')
    if diff:
        sql += " AND diff_levels LIKE ?"
        params.append(f'%{diff}%')
    if q_type:
        sql += " AND q_types LIKE ?"
        params.append(f'%{q_type}%')
    sql += " ORDER BY create_time DESC"
    cur.execute(sql, params)
    rows = cur.fetchall()
    res = []
    for r in rows:
        res.append({
            "id": r[0],
            "title": r[1],
            "full_content": r[2],
            "scene": r[3],
            "organ_systems": json.loads(r[4]),
            "topics": json.loads(r[5]),
            "diff_levels": json.loads(r[6]),
            "q_types": json.loads(r[7]),
            "master_levels": json.loads(r[8]),
            "create_time": r[9]
        })
    conn.close()
    return res

def get_question_by_ids(id_list: List[int]) -> List[Dict]:
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    placeholders = ",".join(["?"] * len(id_list))
    cur.execute(f"SELECT * FROM question_bank WHERE id IN ({placeholders})", id_list)
    rows = cur.fetchall()
    res = []
    for r in rows:
        res.append({
            "id": r[0],
            "title": r[1],
            "full_content": r[2],
            "scene": r[3],
            "organ_systems": json.loads(r[4]),
            "topics": json.loads(r[5]),
            "diff_levels": json.loads(r[6]),
            "q_types": json.loads(r[7]),
            "master_levels": json.loads(r[8]),
            "create_time": r[9]
        })
    conn.close()
    return res

# ---------------- 药物库操作 ----------------
def add_drug(drug_name: str, organ_system: str, drug_class: str):
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    cur.execute("INSERT OR IGNORE INTO drug_list (drug_name, organ_system, drug_class) VALUES (?,?,?)",
                (drug_name, organ_system, drug_class))
    conn.commit()
    conn.close()

def get_all_drugs() -> List[Dict]:
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    cur.execute("SELECT * FROM drug_list")
    rows = cur.fetchall()
    res = []
    for r in rows:
        res.append({"id": r[0], "drug_name": r[1], "organ_system": r[2], "drug_class": r[3]})
    conn.close()
    return res

# ---------------- Prompt模板操作 ----------------
def save_prompt_template(tpl_name: str, content: str, model_type: str):
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    cur.execute('''
    INSERT OR REPLACE INTO prompt_template (template_name, prompt_content, model_type, create_time)
    VALUES (?,?,?,?)
    ''', (tpl_name, content, model_type, now))
    conn.commit()
    conn.close()

def get_all_templates() -> List[Dict]:
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    cur.execute("SELECT * FROM prompt_template ORDER BY create_time DESC")
    rows = cur.fetchall()
    res = []
    for r in rows:
        res.append({
            "id": r[0],
            "template_name": r[1],
            "prompt_content": r[2],
            "model_type": r[3],
            "create_time": r[4]
        })
    conn.close()
    return res

# 初始化数据库
init_db()
# 预置基础药物
base_drugs = [
    ("阿莫西林", "抗感染系统", "青霉素类"),
    ("头孢曲松", "抗感染系统", "头孢三代"),
    ("氯沙坦", "心血管系统", "ARB降压药"),
    ("二甲双胍", "内分泌/糖尿病", "双胍类降糖药"),
    ("阿托伐他汀", "心血管系统", "他汀降脂"),
    ("呋塞米", "肾脏泌尿", "袢利尿剂")
]
for dname, sys, cls in base_drugs:
    add_drug(dname, sys, cls)