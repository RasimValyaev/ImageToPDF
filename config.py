import os
import sys

if os.environ['COMPUTERNAME'] == 'PRESTIGEPRODUCT':
    CONFIG_PATH = r"d:\Prestige\Python\Config"
    CONFIG_PATH_NOVAPOSHTA = r"D:\Prestige\Python\NovaPoshta"
else:
    CONFIG_PATH = r"C:\Rasim\Python\Config"
    CONFIG_PATH_NOVAPOSHTA = r"c:\Rasim\Python\NovaPoshta"

sys.path.append(os.path.abspath(CONFIG_PATH))
sys.path.append(os.path.abspath(CONFIG_PATH_NOVAPOSHTA))