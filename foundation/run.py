import execjs
import numpy as np
from matplotlib import pyplot as plt
import foundation_function


foundation_index_list = [2]  # 需要查询的基金行号，从第二行开始


if __name__ == '__main__':
    foundation_function.update_foundation_info(foundation_index_list)
    foundation_function.compare_transaction_point(foundation_index_list)
