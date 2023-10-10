import os
base_directory = 'AAA'
courtier = 'ASCOMA - SN'
dirs = os.path.join(base_directory,courtier)
try:
    if not os.path.exists(dirs):
        os.makedirs(dirs)
except Exception as e:
    print(e)
print(str(dirs).replace("\\","/"))