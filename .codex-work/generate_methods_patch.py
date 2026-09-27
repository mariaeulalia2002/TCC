import json, re
from pathlib import Path

P = json.loads(Path('.codex-work/docx-paragraphs.json').read_text(encoding='utf-8'))
P = [p for p in P if 396 <= p['index'] <= 597 and p['text'].strip()]

def esc(s):
    for a,b in [('\\',r'\textbackslash{}'),('&',r'\&'),('%',r'\%'),('#',r'\#'),('_',r'\_'),('{',r'\{'),('}',r'\}')]: s=s.replace(a,b)
    return s.replace('–','--').replace('—','---').replace('×',r'$\times$').replace('µ',r'\textmu{}').replace('°',r'$^\circ$')

heads = {
 '4.1   Sísmica':('section','Sísmica'), '4.1.1 Reconhecimento de depósitos e complexos de transporte de massa':('subsection','Reconhecimento de depósitos e complexos de transporte de massa'),
 '4.1.1.1 Cosseno de fase instantânea':('subsubsection','Cosseno de fase instantânea'), '4.1.2 Geração de superfícies sísmicas e análise geomorfométrica':('subsection','Geração de superfícies sísmicas e análise geomorfométrica'),
 '4.2 Caracterização morfométrica dos depósitos':('section','Caracterização morfométrica dos depósitos'),
}
subs=['Preparação dos horizontes e geração das superfícies','Cálculo das isócronas','Área, volume e perímetro','Definição dos marcos operacionais para o cálculo dos comprimentos','Construção das linhas dos limites e definição do eixo no QGIS','Comprimentos total, evacuado e do depósito','Altura da zona evacuada e altura total de queda','Altura e inclinação da cicatriz','Inclinação do declive não deformado figura de exemplo','Inclinação à frente do pé do depósito figura de exemplo','Conversão tempo–profundidade das medidas morfométricas','Incertezas nas medidas morfométricas']
for i,n in enumerate(subs,1): heads[f'4.{i} {n}']=('subsection',n.replace(' figura de exemplo','').replace('–','--'))
subsubs=['Obtenção das velocidades nos poços','Ponderação pela distância aos depósitos','Conversão dos intervalos sedimentares','Procedimentos que envolveram a coluna d’água','Limitações da conversão']
for i,n in enumerate(subsubs,1): heads[('4. 11. 2 ' if i==2 else f'4.11.{i} ')+n]=('subsubsection',n.replace('’',"'"))
unc=['Resolução vertical e conteúdo espectral do dado sísmico','Repetibilidade do picking','Incerteza associada à interpolação','Resolução espacial e medidas planimétricas','Conversão tempo–profundidade','Propagação da incerteza para os parâmetros morfométricos']
for i,n in enumerate(unc,1): heads[f'4.12.{i} {n}']=('subsubsection',n.replace('–','--'))

eq = {
'Δt(x,y)=ttopo(x,y)-tbase(x,y)':(r'\Delta t(x,y)=t_{\mathrm{topo}}(x,y)-t_{\mathrm{base}}(x,y)','Espessura temporal da isócrona'),
'Lt=LA-C':(r'L_t=L_{A-C}','Comprimento total do movimento de massa'),'Ld=LB-C':(r'L_d=L_{B-C}','Comprimento do depósito'),'Le=LA-B=Lt-Ld':(r'L_e=L_{A-B}=L_t-L_d','Comprimento da zona evacuada'),
'Δte=tB-tA e Δtt=tC-tA':(r'\Delta t_e=t_B-t_A \quad\text{e}\quad \Delta t_t=t_C-t_A','Intervalos temporais'),'He=|zB-zA|':(r'H_e=\lvert z_B-z_A\rvert','Altura da zona evacuada'),'Ht=|zC-zA|':(r'H_t=\lvert z_C-z_A\rvert','Altura total de queda'),'Δzd=|zC-zB|':(r'\Delta z_d=\lvert z_C-z_B\rvert','Desnível distal'),'Hs=zBs-zAs':(r'H_s=z_{B_s}-z_{A_s}','Altura da cicatriz'),
'Ss=arctan|zQ-zP|LP-Q180π':(r'S_s=\arctan\left(\frac{\lvert z_Q-z_P\rvert}{L_{P-Q}}\right)\frac{180}{\pi}','Inclinação máxima da cicatriz'),'S=arctan|zP2-zP1|LP1-P2180π':(r'S=\arctan\left(\frac{\lvert z_{P2}-z_{P1}\rvert}{L_{P1-P2}}\right)\frac{180}{\pi}','Inclinação do declive'),'St=arctan|zP4-zP3|LP3-P4180π':(r'S_t=\arctan\left(\frac{\lvert z_{P4}-z_{P3}\rvert}{L_{P3-P4}}\right)\frac{180}{\pi}','Inclinação à frente do pé'),
'vint=2000(z2-z1)t2-t1':(r'v_{\mathrm{int}}=\frac{2000(z_2-z_1)}{t_2-t_1}','Velocidade intervalar'),'h=(ttopo-tbase)vMTD2000':(r'h=\frac{(t_{\mathrm{topo}}-t_{\mathrm{base}})v_{\mathrm{MTD}}}{2000}','Conversão da espessura'),'Δz=Δt15002000':(r'\Delta z=\frac{\Delta t\,1500}{2000}',"Conversão na coluna d'água"),'zP=tfundo,P1500+(tP-tfundo,P)vMTD2000':(r'z_P=\frac{t_{\mathrm{fundo},P}1500+(t_P-t_{\mathrm{fundo},P})v_{\mathrm{MTD}}}{2000}','Profundidade de ponto soterrado'),
'RTWT=12fdom':(r'R_{\mathrm{TWT}}=\frac{1}{2f_{\mathrm{dom}}}','Resolução vertical temporal'),'RTWT=500fdom':(r'R_{\mathrm{TWT}}=\frac{500}{f_{\mathrm{dom}}}','Resolução vertical em milissegundos'),'di=t2,i-t1,i':(r'd_i=t_{2,i}-t_{1,i}','Diferença entre interpretações'),'upick=sd2':(r'u_{\mathrm{pick}}=\frac{s_d}{\sqrt{2}}','Incerteza de repetibilidade'),'uinterp≈RMSE':(r'u_{\mathrm{interp}}\approx\mathrm{RMSE}','Incerteza de interpolação'),
'v=wiviwi,wi=1di':(r'\bar v=\frac{\sum_i w_i v_i}{\sum_i w_i},\quad w_i=\frac{1}{d_i}','Média ponderada da velocidade'),'v±uv':(r'\bar v\pm u_v','Velocidade com incerteza'),'uF=∂F∂x1ux12+∂F∂x2ux22+…':(r'u_F=\sqrt{\left(\frac{\partial F}{\partial x_1}u_{x_1}\right)^2+\left(\frac{\partial F}{\partial x_2}u_{x_2}\right)^2+\cdots}','Propagação geral de incertezas'),'ut=upick2+uinterp2':(r'u_t=\sqrt{u_{\mathrm{pick}}^2+u_{\mathrm{interp}}^2}','Incerteza temporal'),'uΔt=ut12+ut22':(r'u_{\Delta t}=\sqrt{u_{t_1}^2+u_{t_2}^2}','Incerteza da diferença temporal'),'z=vt2':(r'z=\frac{vt}{2}','Conversão de tempo em profundidade'),'uzz=uvv2+utt2':(r'\frac{u_z}{z}=\sqrt{\left(\frac{u_v}{v}\right)^2+\left(\frac{u_t}{t}\right)^2}','Incerteza relativa da profundidade'),'uF=i∂F∂xiuxi2':(r'u_F=\sqrt{\sum_i\left(\frac{\partial F}{\partial x_i}u_{x_i}\right)^2}','Formulação geral para variáveis independentes')}

cites={'(Tearpock & Bischke, 2003)':r'\cite{tearpock2003}','(Alvarenga, 2016)':r'\cite{alvarenga2016}','(Preissler, 2016)':r'\cite{preissler2016}','(Davis, 2002)':r'\cite{davis2002}','(Riley et al., 1999; Wilson et al., 2007)':r'\cite{riley1999,wilson2007}','(QGIS Development Team, 2025)':r'\cite{qgis2025}','(Mackenzie, 1981)':r'\cite{mackenzie1981}','(Kallweit e Wood, 1982)':r'\cite{kallweit1982}','(Isaaks e Srivastava, 1989)':r'\cite{isaaks1989}','(Shepard, 1968)':r'\cite{shepard1968}','Taylor (1977)':r'\citeauthoronline{taylor1997} (\citeyear{taylor1997})','Taylor (1997)':r'\citeauthoronline{taylor1997} (\citeyear{taylor1997})','Savitzky–Golay (1964)':r'\citeauthoronline{savitzky1964} (\citeyear{savitzky1964})','Clare et al. (2018)':r'\citeauthoronline{clare2018} (\citeyear{clare2018})','(Clare et al., 2018)':r'\cite{clare2018}'}
ital=['Petrel®','Surfer®','QGIS®','Python®','Post-Stack Time Migration','binary header','trace headers','bin size','Manual Interpretation','loop tie','inlines','crosslines','mis-ties','pockmarks','Guided Autotracking','two-way travel time','grid','raster','Make Surface','Surfer Binary Grid','Operations','Volume Below Surface','Profile Tool','picking']

def prose(s):
    s=esc(s)
    for a,b in cites.items(): s=s.replace(esc(a),b)
    for w in ital: s=s.replace(esc(w),r'\textit{'+esc(w)+'}')
    s=re.sub(r'Figura [Xx]',r'\\TextoModelo{Figura a inserir}',s)
    s=re.sub(r'Anexo [Xx1]',r'\\TextoModelo{Anexo a definir}',s)
    return s

out=[r'\chapter{Materiais e métodos}\label{cap:metodos}','']
for p in P:
    s=p['text'].strip()
    if s in heads:
        cmd,title=heads[s]; out += [f'\\{cmd}{{{title}}}','']; continue
    if s.startswith('Figura 13 —'):
        out += [r'A representação esquemática dos parâmetros morfométricos é apresentada na Figura~\ref{fig:parametros-morfometricos}.','']; continue
    if s.startswith('Fonte: Traduzido de Clare'): continue
    if s.startswith(r'\[ \overline{v}'):
        out += [r'\begin{equation}',r'\overline{v}_{\mathrm{MTD}}=\frac{\sum_i v_i/d_i}{\sum_i 1/d_i}',r'\end{equation}',r'\equationdescription{Velocidade ponderada do MTD}','']; continue
    if s.startswith(r'\[ \boxed{u_{pick}'):
        out += [r'\begin{equation}',r'\boxed{u_{\mathrm{pick}}\approx 0{,}4\,\mathrm{ms}}',r'\end{equation}',r'\equationdescription{Resultado da incerteza de picking}','']; continue
    if s in eq:
        formula,desc=eq[s]; out += [r'\begin{equation}',formula,r'\end{equation}',f'\\equationdescription{{{desc}}}','']; continue
    t=prose(s)
    if s.startswith('Observação:') or 'Colocar figura' in s or 'figura a inserir' in s.lower(): t=r'\noindent\colorbox{yellow}{\parbox{0.94\linewidth}{'+t+'}}'
    out += [t,'']

content='\n'.join(out)
print('*** Begin Patch')
print('*** Delete File: latex/capitulos/04-materiais-metodos.tex')
print('*** Add File: latex/capitulos/04-materiais-metodos.tex')
for line in content.splitlines(): print('+'+line)
print('*** End Patch')
