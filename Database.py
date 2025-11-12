# Database.py
import oracledb
from contextlib import contextmanager

DB_CONFIG = {
    'host': '*',
    'port': '*',
    'service_name': '*',
    'user': '*',
    'password': '*'
}

POOL = None

def get_pool():
    global POOL
    if POOL is None:
        POOL = oracledb.SessionPool(
            user=DB_CONFIG['user'],
            password=DB_CONFIG['password'],
            dsn=oracledb.makedsn(DB_CONFIG['host'], DB_CONFIG['port'], service_name=DB_CONFIG['service_name']),
            min=2, max=10, increment=1, threaded=True
        )
    return POOL

@contextmanager
def get_connection():
    pool = get_pool()
    conn = pool.acquire()
    try:
        yield conn
    finally:
        pool.release(conn)
