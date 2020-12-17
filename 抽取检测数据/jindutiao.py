import time
from tqdm import tqdm

with tqdm(total=100) as pbar:
    num = 1
    for i in range(58):
        time.sleep(0.1)
        pbar.update(100/58)
