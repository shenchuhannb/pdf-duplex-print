# -*- coding: utf-8 -*-
"""pytest 共享配置：把项目根目录加入 sys.path（tests 在子目录中运行）。"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
