import sys
sys.setrecursionlimit(sys.getrecursionlimit() * 5) # Or a higher value if needed

from scripts.modelo_cisco import processar_mod_cisco

if __name__ == "__main__":
    processar_mod_cisco()
