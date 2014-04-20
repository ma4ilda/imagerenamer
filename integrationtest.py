from ImageRenamer import *
import sys, os
def main(args):
    vars = {}
    url = []
    for a in args:
        if '=' in a:
            d = a.split("=")
            vars[d[0]] = d[1]
        else:
            if os.path.exists(a):
                url.append(a)
            else:
                print a + " doesn't exist.\n"
                sys.exit(2)
    helper = ConfigHelper("ImageRenamer")
    print "Tasks :"
    for t in helper.read_config():
        print t
    for u in url:
        excel_reader = ExcelReader(u, helper, var_dict = vars)
        excel_reader.read_sheet(0)
    
if __name__ == '__main__':
    main(sys.argv[1:])