import datetime

from libs.login import login_cookies
import time

t1 = time.time()
login_cookies()
print(datetime.timedelta(seconds=time.time() - t1))
print("=" * 10)
