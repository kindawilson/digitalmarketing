from pathlib import Path
from html import escape
from urllib.parse import quote
from docx import Document
from docx.oxml.ns import qn
import re

ROOT = Path(r"C:\Users\kinda\Documents\Codex\digitalmarketing\Week 3")
DOCX = Path(r"C:\Users\kinda\AppData\Local\Temp\browser-use\exports\Website - UX _ UI-e2feb425-2549-40bb-8f8e-3b4de72eb4bd.docx")
BASE = "https://kindawilson.github.io/digitalmarketing/Week%203/"
doc = Document(DOCX)
P = doc.paragraphs

# Reuse the already-approved Week 3/Week 1-2 visual system without changing it.
sample = (ROOT / "Week 3 Section 1 Human-Centered Design and Short-Term Memory.html").read_text(encoding="utf-8")
css = re.search(r"<style>(.*?)</style>", sample, re.S).group(1)

def linkify(text):
    escaped = escape(text).replace('\n', '<br>')
    return re.sub(r'(https?://[^\s<]+)', lambda m: f'<a href="{m.group(1).rstrip(")")}" target="_blank" rel="noopener noreferrer">{m.group(1).rstrip(")")}</a>{")" if m.group(1).endswith(")") else ""}', escaped)

def run_html(run):
    if not run.text:
        return ''
    out = linkify(run.text)
    if run.bold: out = f'<strong>{out}</strong>'
    if run.italic: out = f'<em>{out}</em>'
    if run.underline: out = f'<u>{out}</u>'
    return out

def rich(i):
    p = P[i]
    source=p.text
    pieces=[]; pos=0
    for run in p.runs:
        if not run.text: continue
        found=source.find(run.text,pos)
        if found < 0: continue
        if found > pos: pieces.append(linkify(source[pos:found]))
        pieces.append(run_html(run)); pos=found+len(run.text)
    if pos < len(source): pieces.append(linkify(source[pos:]))
    return ''.join(pieces).strip() or linkify(source.strip())

def plain(i): return P[i].text.strip()

def img(name, alt, caption=''):
    cap=f'<figcaption class="caption">{caption}</figcaption>' if caption else ''
    return f'<figure class="image-card"><img src="{BASE}{quote(name)}" alt="{escape(alt)}">{cap}</figure>'

def frame(name, title, height, cls=''):
    return f'<div class="embed-card"><iframe class="section-iframe {cls}" src="{BASE}{quote(name)}" title="{escape(title)}" width="100%" height="{height}" scrolling="no" loading="lazy" allow="fullscreen" style="display:block;width:100%;border:0"></iframe></div>'

def clean_bullet(i):
    s=rich(i)
    s=re.sub(r'^(?:·|•|▪|–|—|\s|&nbsp;)+', '', s).strip()
    return s

def clean_number(i):
    return re.sub(r'^\d+\.\s*', '', rich(i)).strip()

def render(start, end, headings=(), h3s=(), list_ranges=(), specials=None, skip=()):
    specials=specials or {}; headings=set(headings); h3s=set(h3s); skip=set(skip)
    list_start={a:(a,b,tag) for a,b,tag in list_ranges}; covered={i for a,b,_ in list_ranges for i in range(a,b+1)}
    chunks=[]; card=[]; in_excerpt=False
    def flush():
        nonlocal card
        if card:
            chunks.append('<article class="content-card">'+''.join(card)+'</article>'); card=[]
    i=start
    while i<=end:
        if i in specials:
            card.append(specials[i]); i+=1; continue
        if i in skip or not plain(i): i+=1; continue
        if i in headings:
            flush(); card.append(f'<h2>{rich(i)}</h2>'); i+=1; continue
        if i in h3s:
            card.append(f'<h3>{rich(i)}</h3>'); i+=1; continue
        if i in list_start:
            _,b,tag=list_start[i]; items=''.join(f'<li>{clean_number(j) if tag=="ol" else clean_bullet(j)}</li>' for j in range(i,b+1) if plain(j))
            card.append(f'<{tag}>{items}</{tag}>'); i=b+1; continue
        if i in covered: i+=1; continue
        low=plain(i).lower()
        if low.startswith('<excerpt from:'):
            label=plain(i).strip('<>')
            label=re.sub(r'^Excerpt from:\s*', 'Excerpt from ', label, flags=re.I)
            card.append(f'<div class="excerpt"><span class="excerpt-label">{escape(label)}</span>'); in_excerpt=True; i+=1; continue
        if low.startswith('<end excerpt'):
            card.append('</div>'); in_excerpt=False; i+=1; continue
        card.append(f'<p>{rich(i)}</p>'); i+=1
    if in_excerpt: card.append('</div>')
    flush(); return ''.join(chunks)

footer='''<footer class="footer mceNonEditable" style="border-top:7px solid #ddaf4b;margin-top:28px;padding:30px 24px;background:#4d5259;color:#ffffff;text-align:center;"><div class="footer-inner" style="width:min(100%,1160px);margin:0 auto;font-size:14px;"><img src="/shared/HTML-Template-Library/HTML-Templates-V5/img/Dimensional_V_White_Lockup.png" alt="Vanderbilt University" style="max-width:260px;width:100%;height:auto;"></div></footer>'''
script='''<script>document.querySelectorAll('.tooltip-button').forEach(function(button){button.addEventListener('click',function(){var wrap=button.closest('.tooltip-wrap');var open=!wrap.classList.contains('open');document.querySelectorAll('.tooltip-wrap.open').forEach(function(item){item.classList.remove('open');item.querySelector('button').setAttribute('aria-expanded','false')});wrap.classList.toggle('open',open);button.setAttribute('aria-expanded',String(open))})});</script>'''

def write_page(filename, number, title, subtitle, takeaways, body, show_takeaways=True):
    take=''
    if show_takeaways:
        take='<aside class="takeaways"><h2>Key Takeaways</h2><ul>'+''.join(f'<li>{x}</li>' for x in takeaways)+'</ul></aside>'
    full=BASE+quote(filename)
    html=f'''<!DOCTYPE html>\r\n<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>{escape(filename[:-5])}</title><link rel="stylesheet" href="https://s.brightspace.com/lib/fonts/0.6.1/fonts.css"><link rel="stylesheet" href="https://templates.lcs.brightspace.com/lib/assets/css/styles.min.css"><link rel="stylesheet" href="/d2l/le/contentstyler/6606/files/View" data-override="override"><style>{css}</style></head><body><main class="course-page"><header class="course-header" style="position:relative;overflow:hidden;padding:34px 24px 32px;border-bottom:7px solid #ddaf4b;background:#0f3f52;color:#fff"><div class="inner"><p class="section-number">{escape(number)}</p><h1>{escape(title)}</h1><p class="subtitle">{escape(subtitle)}</p></div></header><section class="course-content">{take}<p class="full-link"><a href="{full}" target="_blank" rel="noopener noreferrer">Open this lesson full screen</a></p>{body}</section></main>{footer}{script}</body></html>'''
    (ROOT/filename).write_text(html, encoding='utf-8', newline='')

# Overview: preserve the three supplied paragraphs and add only the objectives/resources explicitly requested by the production instructions.
overview=render(2,6)
overview+='''<article class="content-card"><h2>Learning Objectives</h2><ul><li>Explain the difference between user experience and user interface.</li><li>Describe how human-centered design supports digital marketing.</li><li>Apply cognitive-load and visual-design principles to digital materials.</li><li>Evaluate landing-page and mobile-interface choices from the user’s perspective.</li></ul></article><article class="content-card"><h2>Class Preparation</h2><p>Complete Sections 1–5 and interact with the embedded examples.</p></article><section class="resources"><h2>Resources</h2><div class="excerpt"><span class="excerpt-label">Slides and Video</span><p>Course slides and the weekly overview video will be linked here in Brightspace.</p></div></section>'''
write_page('Week 3 Overview.html','Week 3','Website, UX, and UI',plain(2),[],overview,False)

s1=render(18,47,headings=(18,25,36),specials={27:'<div class="excerpt"><span class="excerpt-label">Excerpt from <a href="https://www.amazon.com/Design-Everyday-Things-Revised-Expanded/dp/0465050654" target="_blank" rel="noopener noreferrer"><em>The Design of Everyday Things</em></a></span>',32:'</div>'},skip=(27,32))
write_page('Week 3 Section 1 Human-Centered Design and Short-Term Memory.html','Section 1',plain(10),'Human-centered design starts by accepting how people actually think, remember, and behave.',[rich(13),rich(14)],s1)

s2_special={
77:img('Wall of Text Above the Fold.png','A fake webpage with one very large paragraph filling the area above the fold.'),
79:img('Chunked Content Above the Fold.png','The same information organized into shorter, related chunks.'),
94:img('Von Restorff Visually Distinctive Action.png','An important action made visually distinctive.'),
95:img('Von Restorff Distinctive Purchase Option.png','A recommended purchase option made visually distinctive.'),
92:f'<ol><li>{rich(92)}</li><li>{rich(93)}</li></ol>',
96:f'<ol start="3"><li>{rich(96)}</li></ol>',
117:frame('Website Mental Model Examples.html','Website mental model examples',610,'iframe-mental'),
131:frame('cognitive-load-laws-mockup-popup-v2.html','Combined Laws of UX interaction',760,'iframe-laws')}
s2=render(60,131,headings=(70,86,100,109,127),h3s=(74,91,103,120),list_ranges=((104,107,'ul'),(121,123,'ul')),specials=s2_special,skip=(77,79,82,94,95,117,131))
write_page('Week 3 Section 2 Short-Term Memory and Related Strategies.html','Section 2',plain(51),'Knowing what we do about short-term memory, we can design digital marketing that is easier to process.',[clean_bullet(i) for i in range(54,58)],s2)

s3_special={152:'<div class="action-list">'+''.join(f'<p>{rich(j)}</p>' for j in range(152,158))+'</div>',161:f'<div class="excerpt"><span class="excerpt-label">{rich(161)}</span></div>',167:frame('alignment examples.html','Alignment examples',545,'iframe-carousel'),182:frame('hierarchy examples.html','Hierarchy examples',545,'iframe-carousel'),192:f'<p>{rich(192)}</p>'+img('typography-hierarchy-example.png','An example showing visual hierarchy through typography.',rich(193)),200:frame('white space examples.html','White space examples',545,'iframe-carousel'),223:img('poor-color-contrast-example.jpg','An example with poor color contrast.')}
s3=render(144,223,headings=(148,163,172,198,206),h3s=(210,),list_ranges=((212,219,'ul'),),specials=s3_special,skip=(153,154,155,156,157,161,167,182,192,193,200,204,221,223))
write_page('Week 3 Section 3 Design Principles and Layout Terms.html','Section 3',plain(135),'Apply foundational design principles to websites, ads, emails, and other digital materials.',[clean_bullet(139),clean_bullet(140)],s3)

tip='''<span class="tooltip-wrap"><button class="tooltip-button" type="button" aria-label="What does above the fold mean?" aria-expanded="false">?</button><span class="tooltip-box" role="tooltip">Above the fold refers to information that is showing on the screen when you land on a page. Much like the information that would show on the front page of a news paper when you “fold” it.</span></span>'''
s4_special={231:frame('website image carousel.html','Website image carousel',620,'iframe-carousel'),250:img('landing-page-three-questions.png','A visual showing the three questions a homepage or landing page needs to answer.'),256:f'<p>For the design “above the fold” {tip} (Remember KISS - Keep It Simple, Stupid):</p>',267:img('landing-page-template-example.png','An example of the recommended above-the-fold page template.'),271:frame('landing-page-elements-break-it-standalone-iframe.html','Break the landing page interaction',780,'iframe-break'),274:f'<div class="geek-out">{rich(274)}</div>'}
s4=render(229,274,headings=(242,254),h3s=(246,),list_ranges=((247,249,'ul'),(257,263,'ul')),specials=s4_special,skip=(231,250,256,267,271,274))
write_page('Week 3 Section 4 Putting It All Together A Page Template.html','Section 4',plain(225),'Combine the Laws of UX and design principles into a practical landing-page template.',['A landing page or homepage should quickly answer who you are, why customers should buy, and what to do next.','Use a compelling image, short headline, supporting information, clear call to action, and whitespace.'],s4)

s5_special={298:img('mobile-holding-positions.png','Common ways people hold and touch a mobile phone.',rich(298)),308:img('mobile-touch-accuracy.jpg','A chart showing touch accuracy for specific parts of a mobile screen.',rich(308)),323:img('touch-friendly-information-design.png','A touch-friendly information-design framework.',rich(323))}
s5=render(285,323,headings=(290,313),specials=s5_special,skip=(287,298,308,323))
s5=s5.replace('<article class="content-card">',f'<article class="content-card"><div class="excerpt"><span class="excerpt-label">Excerpt from <a href="https://www.uxmatters.com/mt/archives/2017/03/design-for-fingers-touch-and-people-part-1.php" target="_blank" rel="noopener noreferrer">Design for Fingers, Touch, and People: Part 1</a></span>',1)
s5=s5.replace('</article>','</div></article>',1)
write_page('Week 3 Section 5 Mobile Interaction.html','Section 5',plain(276),'Design mobile experiences for real hands, changing grips, and imprecise touch.',[clean_bullet(i) for i in range(279,283)],s5)

print('Restored source wording in all six Week 3 lesson pages.')
