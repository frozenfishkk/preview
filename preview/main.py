import sys
from init import Init
import argparse

# 按装订区域中的绿色按钮以运行脚本。
if __name__ == '__main__':

    parser = argparse.ArgumentParser(description='')

    parser.add_argument('--pathdamage', type=str, default=None)
    parser.add_argument('--current_excel', type=str, default="default")
    args = parser.parse_args()
    window = Init(args.pathdamage, args.current_excel)
    window.show()
    sys.exit(Init.app.exec())



