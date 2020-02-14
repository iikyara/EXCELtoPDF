DEBUG = False
def log(*arg, **darg):
    if DEBUG:
        print(*arg, **darg)
