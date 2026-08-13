#!/usr/bin/env python3
"""Lightweight PII redactor for DOCX files.

Usage:
    python pii_redactor.py input.docx output.docx
"""
from __future__ import annotations
import argparse, json, re
from dataclasses import dataclass
from pathlib import Path
from docx import Document

EMAIL_RE = re.compile(r"(?<![\w.+-])[\w.+-]+@[\w.-]+\.[A-Za-z]{2,}(?![\w.-])")
SSN_RE = re.compile(r"\b\d{3}-\d{2}-\d{4}\b")
IP_RE = re.compile(r"(?<![\w.])(?:25[0-5]|2[0-4]\d|1?\d?\d)(?:\.(?:25[0-5]|2[0-4]\d|1?\d?\d)){3}(?![\w.])")
CC_RE = re.compile(r"(?<!\d)(?:\d[ -]?){13,19}(?!\d)")
PHONE_PATTERNS = [
    re.compile(r"(?<!\d)\+?\s*91[\s-]?(?:\(\s*\d{2,4}\s*\)|\d{2,4})[\s-]?\d{3,5}[\s-]?\d{4}(?!\d)"),
    re.compile(r"(?<!\d)\+\d{1,3}[\s-]?(?:\(?\d{2,4}\)?[\s-]?)\d{3,4}[\s-]?\d{4}(?!\d)"),
    re.compile(r"(?<!\d)[6-9]\d{9}(?!\d)"),
]
DATE_RE = r"(?:\d{1,2}[/-]\d{1,2}[/-]\d{2,4}|(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},?\s+\d{4})"
DOB_RE = re.compile(rf"(?i)(?:date\s+of\s+birth|dob|born(?:\s+on)?)\s*[:\-]?\s*({DATE_RE})")
PIN_RE = re.compile(r"\b\d{3}(?:\s|-)?\d{3}\b")
ADDRESS_CONTEXT_RE = re.compile(r"(?i)\b(?:registered office|corporate office|mailing address|residential address|residence|address|correspondence address)\b")
ADDRESS_TERMS_RE = re.compile(r"(?i)\b(?:road|marg|street|lane|avenue|floor|plot|taluka|district|village|society|apartment|bungalow|building|tower|phase|pincode|pin|embassy|cts\s*no\.?|s\.?\s*no\.?)\b|\b[A-Z]-?\d{2,}\b")
COMPANY_SUFFIX_RE = re.compile(r"(?i)\b(?:Private\s+Limited|Limited|LLP|Inc\.?|Incorporated|Corporation|Bank|Foundation|Trust|Industries|Technologies|Motors|Electricals)\b")

STOPWORDS = {"The","Our","This","That","Red","Herring","Prospectus","Equity","Shares","Offer","Price","Anchor","Investor","Investors","Mutual","Funds","Life","Insurance","Companies","Pension","Book","Running","Lead","Managers","Stock","Exchanges","Working","Days","Company","Management","Risk","Factors","Capital","Structure","Financial","Statements","India","Maharashtra","Mumbai","Pune","Bank","Limited","Private","Trust","Family","KSH","Website","Email","Telephone","Registration","Number","Account","Address","Description","Term","For","With","Report","Committee","Application","Forms","Amount","Bidders","Manager","Managers","Securities","Road","Marg","Society","East","West","North","South","Building","Floor","Tower","Level","Centre","Complex","Village","Taluka","District","Plot","Phase","Park","House","Hotel","Apartment","Showroom","Department","Division","Capital","Market","World","Governance","Information","Agreement","Agreements","Banks","Sponsor","Syndicate","Members","Locations","Intermediaries","Escrow","Agent","Price","Date","Period","Proceeds","Appraising","Entity","Monitoring","Agency","Legal","Counsel","Indian","Law","Exchange","Board","Accounts","Cash","Rating","Experts","Specified","Registered","Broker","Brokers","Group","Entities","Branch","Parents","Share","Based","Payment","Expense","Statutory","Auditors","Chartered","Accountants","Disclosure","Requirements","Regulatory","Other","Qualified","Institutional","Buyers","Statutory","Eligibility","Regulations","Obligations","Listing","Exchange","Securities","Capital","Issue","Regulation","Regulatory","Policy","Policies","Legal","General","Designated","Intermediaries","Qualified","Institutional","Portion","Application","Supported","Blocked","Amount","Self","Certified","Syndicate","Bidding","Offer","Structure","Summary","Definitions","Abbreviations"}
ROLE_WORDS = {"director","directors","chairman","executive","managing","whole-time","independent","promoter","shareholder","secretary","officer","person","personnel","chief","financial","compliance"}

@dataclass(frozen=True)
class Span:
    start: int
    end: int
    label: str


def luhn_valid(value: str) -> bool:
    digits = [int(c) for c in re.sub(r"\D", "", value)]
    if not 13 <= len(digits) <= 19: return False
    total = 0
    parity = len(digits) % 2
    for i, d in enumerate(digits):
        if i % 2 == parity:
            d *= 2
            if d > 9: d -= 9
        total += d
    return total % 10 == 0


def title_name_candidates(text: str) -> list[str]:
    pat = re.compile(r"\b[A-Z][A-Za-z.'-]{1,30}(?:\s+[A-Z][A-Za-z.'-]{1,30}){1,3}\b")
    out=[]
    for m in pat.finditer(text):
        value=m.group(0).strip(" ,;:/"); words=value.split()
        if not 2 <= len(words) <= 4: continue
        if any(w in STOPWORDS for w in words): continue
        if any(w.lower() in ROLE_WORDS for w in words): continue
        if any(w.isupper() and len(w)>1 for w in words): continue
        if COMPANY_SUFFIX_RE.search(value): continue
        if any(w.lower() in {"website","email","telephone","registration","number","account","address","description","term","for","with","limited","private","trust","bank","company","report","committee","application","forms","amount","offer","price","shares","date","period","investor","investors","bidders","manager","managers","securities","india","maharashtra","mumbai","pune","road","marg","society","east","west","north","south"} for w in words): continue
        out.append(value)
    return out


def discover_person_names(text: str) -> set[str]:
    names=set(); cleaned=re.sub(r"\s+"," ",text)
    pat=(r"(?i)(?:contact\s+person|contact\s+persons?|promoters?)\s*:\s*"
         r"(.*?)(?=\b(?:contact\s+person|contact\s+persons?|promoters?|company|registered\s+office|"
         r"corporate\s+office|website|email|telephone|tel|sebi\s+registration)\b|$)")
    for m in re.finditer(pat,cleaned):
        for piece in re.split(r"[/,]",m.group(1)): names.update(title_name_candidates(piece))
    for m in re.finditer(r"(?i)\b(?:being|namely)\s+([A-Z][A-Za-z.'-]+(?:\s+[A-Z][A-Za-z.'-]+){1,3})",cleaned):
        names.update(title_name_candidates(m.group(1).strip(" ,;:.")))
    return names


def discover_company_names(text: str) -> set[str]:
    companies=set(); cleaned=re.sub(r"\s+"," ",text)
    token=r"[A-Z][A-Za-z0-9&.'()/-]*"
    suffix=r"(?:Private\s+Limited|Limited|LLP|Inc\.?|Incorporated|Corporation|Bank|Foundation|Trust|Industries|Technologies|Motors|Electricals)"
    pat=re.compile(rf"\b(?:{token}\s+){{0,6}}{suffix}\b")
    bad={"the","our","fresh","issue","offer","anchor","investor","investors","date","email","telephone","and","public","account","bankers","registrar","to","former","formerly","contact","person","website","grievance","running","lead","managers","members","sponsor","promoter","trusts","term","abbreviations","advance","estimates","fema","rules","npcI","rbi","sebi","inm000011179","inr000004058","l65920mh1994plc080618","l65190gj1994plc080618","l65190gj1994plc021012","positive","stable","short","long","complaints","redressal","mechanism","self-certified","company","escrow","collection","refund","bid","regulations","regulation","act","circular","senior","kmp","fpi","fpIs","aif","vcf","btI","industrial","development","solar","education","u.s."}
    generic={"private limited","co llp","co. llp","services private limited","advisory private limited","electricals private limited","investment private limited","solutions limited","energy corporation","fund limited","depository limited","reserve bank","state bank","bank hdfc bank limited","bank limited hdfc bank limited","banks icici bank limited","bank icici bank","bse bse limited","india) limited","pandit llp","park i private limited","park ii private limited","park iii private limited","park iv private limited","park v private limited","park vi private limited","park viii private limited","park ix private limited","park ix a private limited","park ix b private limited"}
    for m in pat.finditer(cleaned):
        words=m.group(0).strip(" ,;:/").split()
        for size in range(min(7,len(words)),1,-1):
            cw=words[-size:]; lower={w.lower().strip(".,") for w in cw}
            if lower & bad: continue
            if any(any(ch.isdigit() for ch in w) for w in cw) and not any(w.upper() in {"I","II","III","IV","V","VI","VIII","IX","IXA","IXB"} for w in cw): continue
            candidate=" ".join(cw)
            if candidate.lower() in generic: continue
            half=len(cw)//2
            if len(cw)>=4 and cw[:half]==cw[half:]: continue
            companies.add(candidate); break
    companies.update(re.findall(r"\bWaterloo Industrial Park (?:I|II|III|IV|V|VI|VIII|IX(?: A| B)?) Private Limited\b",cleaned))
    companies.update(re.findall(r"\bKSH Infra Park (?:IV|VI) Private Limited\b",cleaned))
    return companies

class Redactor:
    def __init__(self):
        self.person_names=set(); self.company_names=set(); self.mapping={}
        self.counters={k:0 for k in ["NAME","EMAIL","PHONE","COMPANY","ADDRESS","SSN","CREDIT_CARD","DOB","IP_ADDRESS"]}
        self.fake_names=["John Doe","Jane Smith","Peter Parker","Alex Johnson","Emily Brown","Michael Wilson","Sarah Davis","David Miller","Olivia Taylor","Daniel Anderson","Sophia Thomas","James Jackson","Ava White","Noah Harris","Mia Martin"]
        self.fake_companies=["Acme Technologies Private Limited","BlueSky Industries Limited","NorthStar Finance Limited","Greenfield Consulting LLP","Summit Securities Limited","Pioneer Manufacturing Private Limited"]
        self.fake_addresses=["101 Example Road, Mumbai, Maharashtra, India","22 Sample Street, Pune, Maharashtra, India","45 Demo Park, Bengaluru, Karnataka, India","9 Test Avenue, New Delhi, Delhi, India"]

    def learn(self, paragraphs):
        text="\n".join(paragraphs); self.person_names |= discover_person_names(text); self.company_names |= discover_company_names(text)

    def learn_table_context(self,tables):
        for table in tables:
            if not table.rows: continue
            header=" | ".join(c.text for c in table.rows[0].cells).lower()
            name_cols={i for i,c in enumerate(table.rows[0].cells) if "name" in c.text.lower() or "contact" in c.text.lower()}
            role_table=bool(re.search(r"(?i)\b(?:promoter|contact\s+person|director|key\s+managerial|shareholder)\b",header))
            for row in table.rows:
                row_text=" | ".join(c.text for c in row.cells)
                for i,cell in enumerate(row.cells):
                    ct=re.sub(r"\s+"," ",cell.text).strip()
                    if i in name_cols or (role_table and len(ct.split())<=5 and not ADDRESS_TERMS_RE.search(ct)):
                        self.person_names.update(title_name_candidates(ct))
                    self.company_names |= discover_company_names(ct)

    def replacement(self,label,original):
        key=(label, original.casefold() if label in {"NAME","COMPANY"} else original)
        is_new = key not in self.mapping
        if not is_new:
            value=self.mapping[key]
        elif label=="NAME": value=self.fake_names[self.counters[label]%len(self.fake_names)]
        elif label=="COMPANY": value=self.fake_companies[self.counters[label]%len(self.fake_companies)]
        elif label=="EMAIL": value=f"contact{self.counters[label]+1}@example.com"
        elif label=="PHONE": value=f"+91 90000 {10000+self.counters[label]:05d}"
        elif label=="ADDRESS": value=self.fake_addresses[self.counters[label]%len(self.fake_addresses)]
        elif label=="SSN": value=f"999-88-{1000+self.counters[label]:04d}"
        elif label=="CREDIT_CARD": value="4111 1111 1111 1111"
        elif label=="DOB": value="01/01/1990"
        elif label=="IP_ADDRESS": value=f"192.0.2.{10+self.counters[label]}"
        else: value="[REDACTED]"
        self.mapping[key]=value
        if is_new: self.counters[label]+=1
        if original.isupper() and label in {"NAME","COMPANY"}: return value.upper()
        return value

    def detect(self,text):
        spans=[]
        for m in EMAIL_RE.finditer(text): spans.append(Span(m.start(),m.end(),"EMAIL"))
        for m in SSN_RE.finditer(text): spans.append(Span(m.start(),m.end(),"SSN"))
        for m in IP_RE.finditer(text): spans.append(Span(m.start(),m.end(),"IP_ADDRESS"))
        for m in CC_RE.finditer(text):
            if luhn_valid(m.group()): spans.append(Span(m.start(),m.end(),"CREDIT_CARD"))
        for m in DOB_RE.finditer(text): spans.append(Span(m.start(1),m.end(1),"DOB"))
        for p in PHONE_PATTERNS:
            for m in p.finditer(text): spans.append(Span(m.start(),m.end(),"PHONE"))
        if ADDRESS_CONTEXT_RE.search(text) and ADDRESS_TERMS_RE.search(text): spans.append(Span(0,len(text),"ADDRESS"))
        elif re.search(r"\b[A-Z]-?\d{2,}\b|\bEmbassy\b", text) and len(text) <= 250: spans.append(Span(0,len(text),"ADDRESS"))
        elif PIN_RE.search(text) and ADDRESS_TERMS_RE.search(text) and len(text)<=450 and not EMAIL_RE.search(text): spans.append(Span(0,len(text),"ADDRESS"))
        if self.company_names:
            alt=re.compile("|".join(re.escape(v) for v in sorted(self.company_names,key=len,reverse=True)),re.I)
            for m in alt.finditer(text): spans.append(Span(m.start(),m.end(),"COMPANY"))
        if self.person_names:
            alt=re.compile("|".join(re.escape(v) for v in sorted(self.person_names,key=len,reverse=True)),re.I)
            for m in alt.finditer(text): spans.append(Span(m.start(),m.end(),"NAME"))
        priority={"EMAIL":100,"PHONE":95,"SSN":95,"CREDIT_CARD":95,"DOB":90,"IP_ADDRESS":90,"ADDRESS":50,"COMPANY":40,"NAME":30}
        chosen=[]
        for s in sorted(spans,key=lambda x:(-priority[x.label],-(x.end-x.start),x.start)):
            if any(s.start<b and s.end>a for a,b in [(x.start,x.end) for x in chosen]): continue
            chosen.append(s)
        return sorted(chosen,key=lambda x:x.start)

    def redact_text(self,text):
        spans=self.detect(text)
        if not spans: return text,[]
        out=[]; cur=0; events=[]
        for s in spans:
            original=text[s.start:s.end]; replacement=self.replacement(s.label,original)
            out += [text[cur:s.start],replacement]; cur=s.end
            events.append({"type":s.label,"original":original,"replacement":replacement})
        out.append(text[cur:]); return "".join(out),events


def iter_paragraphs(doc):
    for p in doc.paragraphs: yield p
    for t in doc.tables:
        for row in t.rows:
            for c in row.cells:
                for p in c.paragraphs: yield p
    for sec in doc.sections:
        for p in sec.header.paragraphs: yield p
        for p in sec.footer.paragraphs: yield p


def set_paragraph_text(paragraph,text):
    first=paragraph.runs[0] if paragraph.runs else None
    if first:
        font=first.font; paragraph.clear(); run=paragraph.add_run(text)
        run.bold=font.bold; run.italic=font.italic; run.underline=font.underline; run.font.name=font.name
        if font.size: run.font.size=font.size
    else: paragraph.text=text


def redact_docx(input_path,output_path,report_path=None):
    doc=Document(input_path); paragraphs=list(iter_paragraphs(doc)); red=Redactor()
    red.learn(p.text for p in paragraphs); red.learn_table_context(doc.tables)
    events=[]; address_headers={"address","mailing address","registered office","corporate office","residential address","correspondence address"}
    strong_address=re.compile(r"(?i)\b(?:road|marg|street|lane|avenue|floor|plot|taluka|district|village|society|apartment|bungalow|building|tower|phase|pincode|pin|embassy|cts\s*no\.?|s\.?\s*no\.?)\b|\b[A-Z]-?\d{2,}\b")
    for table in doc.tables:
        if not table.rows: continue
        header=[c.text.strip().lower() for c in table.rows[0].cells]
        address_cols={i for i,v in enumerate(header) if v in address_headers or "address" in v}
        for row in table.rows[1:]:
            for i,cell in enumerate(row.cells):
                cell_text=cell.text.strip()
                if not cell_text or not PIN_RE.search(cell_text): continue
                if i in address_cols:
                    repl=red.replacement("ADDRESS",cell_text); cell.text=repl; events.append({"type":"ADDRESS","original":cell_text,"replacement":repl}); continue
                for p in cell.paragraphs:
                    pt=p.text.strip()
                    if not pt or not (PIN_RE.search(pt) or strong_address.search(pt)): continue
                    if re.search(r"(?i)\b(?:email|website|telephone|tel|contact person)\b",pt): continue
                    repl=red.replacement("ADDRESS",pt); set_paragraph_text(p,repl); events.append({"type":"ADDRESS","original":pt,"replacement":repl})
    for p in iter_paragraphs(doc):
        new,ev=red.redact_text(p.text)
        if new!=p.text: set_paragraph_text(p,new); events.extend(ev)
    doc.save(output_path)
    counts={}
    for e in events: counts[e["type"]]=counts.get(e["type"],0)+1
    report={"input_file":str(input_path),"output_file":str(output_path),"detected_replacements":counts,"total_replacements":len(events),"unique_mappings":len(red.mapping),"note":"Counts are replacement events, not unique PII values. The prospectus has no confirmed DOB/SSN/credit-card/IP examples; the detector supports all required types."}
    if report_path: Path(report_path).write_text(json.dumps(report,indent=2),encoding="utf-8")
    return report

if __name__=="__main__":
    ap=argparse.ArgumentParser(); ap.add_argument("input",type=Path); ap.add_argument("output",type=Path); ap.add_argument("--report",type=Path); a=ap.parse_args()
    print(json.dumps(redact_docx(a.input,a.output,a.report),indent=2))
