import os
from sys import argv
import win32com.client
from multiprocessing import Pool
import timeit

currpath = os.path.dirname(os.path.abspath(__file__))
ppttoPDF = 32

def convert(f):
    try:
        print(f'Converting: {f}')
        in_file=os.path.join(str(currpath),f)
        powerpoint = win32com.client.Dispatch("Powerpoint.Application")
        pdf = powerpoint.Presentations.Open(in_file, WithWindow=False)
        pdf.SaveAs(os.path.join(currpath,f[:-5]), ppttoPDF) # formatType = 32 for ppt to pdf
        pdf.Close()
        powerpoint.Quit()
        print(f'Pdf created: {f[:-5]}.pdf')
        #os.remove(os.path.join(root,f))
    except Exception as e:
        print(f'Could not convert {f}')
        print(e)
        
def main():
    start_time = timeit.default_timer()
    if len(argv) > 1:
        global currpath
        currpath = argv[1]

    print(f'Current working directory: {currpath}')

    pptx_files = os.listdir(currpath)
    pptx_files = [ppt for ppt in pptx_files if ppt.endswith('.pptx')]

    ln = len(pptx_files)
    prcs = ln if ln <= 12 else 8
    
    with Pool(prcs) as p:
        p.map(convert, pptx_files)
    print(f'Time elapsed: {timeit.default_timer() - start_time}')

if __name__ == '__main__':
    main()