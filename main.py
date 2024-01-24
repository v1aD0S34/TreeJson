import warnings
from generation import generateTrends, generateOpcMap

if __name__ == "__main__":
    warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
    generateTrends()
    generateOpcMap()
