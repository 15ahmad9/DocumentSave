<!-- هذه الأوامر نطبقها في CMD -->

& msedge.exe --remote-debugging-port=9222 --user-data-dir="C:\edge-selenium"
 
 
edge_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
 
python -m pip install selenium requests


<!-- نفتح المتصفح من RUN -->

msedge.exe --remote-debugging-port=9222 --user-data-dir="C:\edge-selenium"

<!-- نطبق هذا الأمر في CMD حسب مكان وجود Python وملف التخزين الكود -->

C:\Users\ahmad\AppData\Local\Programs\Python\Python313\python.exe "C:\Users\ahmad\Downloads\document\Doc.py"


where python
