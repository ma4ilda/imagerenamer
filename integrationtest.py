from ImageRenamer import *
import sys, os
def main(args):
    variables = {}
    urls = []
    for a in args:
        if '=' in a:
            d = a.split("=")
            variables[d[0]] = d[1]
        else:
            if os.path.exists(a):
                urls.append(a)
            else:
                print a + " doesn't exist.\n"
                sys.exit(2)
    worker = Worker()
    print "Tasks :"
    for t in worker.read_config():
        print t
    for url in urls:
        v = variables.copy()
        reader = ExcelReader(url)
        reader.run(0, worker.execute, v)
    
if __name__ == '__main__':
    main(sys.argv[1:])