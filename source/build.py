# Dựng lại index.html từ tpl.html + thư mục assets. Chạy: python3 build.py
import base64,json,os,re
H=os.path.dirname(os.path.abspath(__file__))
t=open(os.path.join(H,'tpl.html'),encoding='utf-8').read()
t=t.replace('18 giây · 1080p · có thoại','18 giây · có thoại')
t=t.replace('<b>12<small>+</small></b><span>dự án tiêu biểu đủ ngành: ngân hàng, y khoa, marketing, đoàn thể</span>','<b>16</b><span>dự án tiêu biểu: ngân hàng, y khoa, marketing, du lịch, kỷ yếu</span>')
t=t.replace('Nhóm sinh viên môn Kinh doanh quốc tế','Nhóm sinh viên (bài thảo luận)')
A=os.path.join(H,'assets')
def u(f,m='image/webp'):return f'data:{m};base64,'+base64.b64encode(open(os.path.join(A,f),'rb').read()).decode()
for k in ['logo','portrait','demo_poster','v1005','vnbd_poster']: t=t.replace('{{%s}}'%k,u(k+'.webp'))
t=t.replace('{{feat}}',u('kl12.webp')).replace('{{demo}}',u('demo.mp4','video/mp4')).replace('{{vnbd}}',u('vnbd.mp4','video/mp4'))
used=set(re.findall(r'"([a-z0-9]+)"',''.join(re.findall(r'imgs:\[([^\]]*)\]',t))))|{f'fb{k}' for k in range(4)}
t=t.replace('{{imgjson}}',json.dumps({k:u(k+'.webp') for k in used}))
t=re.sub(r'<meta name="description"[^>]*>\n','',t,1);t=t.replace('<title>DUY QUANG Portfolio</title>','',1).replace('<meta name="claude:viewport-fit-sides" content="cover">\n','',1)
head=open(os.path.join(H,'head.html'),encoding='utf-8').read()
open(os.path.join(H,'..','index.html'),'w',encoding='utf-8').write(head+t+'\n</body>\n</html>\n')
print('OK')
