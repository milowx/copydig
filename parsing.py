import os, re
import argparse
from pathlib import Path
import mailbox
from bs4 import BeautifulSoup
from email.header import decode_header
import logging
import hashlib
import tempfile, json
from tqdm import tqdm
from pdfminer.high_level import extract_text as extract_pdf_text
import docx

logging.basicConfig(level=logging.INFO)

EMAIL_RE = re.compile(r'\b[\w\.-]+@[\w\.-]+\.\w+\b')
PHONE_RE = re.compile(r'\+?\d[\d\-\s]{6,}\d')
APIKEY_RE = re.compile(r'(AIza[0-9A-Za-z\-_]{35}|[A-Za-z0-9_\-]{32,})')

def sha256_hex(s:str)->str:
    return hashlib.sha256(s.encode("utf-8")).hexdigest()

def sanatize_id(s: str, max_len=180):
    s = str(s)
    s = re.sub(r"[^A-Za-z0-9._-]", "-", s)

    return s[:max_len or sha256_hex(s)[:16]]



def save_doc(outdir: Path, doc: dict, id_field: str="id", allow_overwrite: bool = False, set_permisions: bool=True) -> Path:
    outdir = Path(outdir)
    outdir.mkdir(parents=True, exist_ok=True)

    doc_id = None
    if id_field in doc and doc[id_field]:
        doc_id = sanatize_id(doc[id_field])
    else:
        key_src = (str(doc.get("source", ""))+"|"+
                   str(doc.get("timestamp"))+"|"+
                   str(doc.get("text","")) )[:2000]
        doc_id = sha256_hex(key_src)[:64]

    filename = f"{doc_id}.json"
    target_path = outdir/filename

    if target_path.exists() and not allow_overwrite:
        raise FileExistsError(f"File already exists and overwrite disabled: {target_path}")
    

    fd = None
    tmp_path = None
    try:
        fd, tmp_path = tempfile.mkstemp(prefix=filename + ".", dir=str(outdir))

        with os.fdopen(fd, "w", encoding="utf-8") as tmpf:
            json.dump(doc, tmpf, ensure_ascii=False)
            tmpf.flush()
            os.fsync(tmpf.fileno())
        
        os.replace(tmp_path, str(target_path))
        tmp_path=None

        if set_permisions:
            try:
                os.chmod(target_path, 0o600)
            except Exception:
                logging.debug("Could not set permissions on %s", target_path)
        return target_path
    except Exception as e:
        if tmp_path and os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except Exception:
                logging.debug("Failed to remove tmp file %s", tmp_path)
        raise


def remove_pii(texto: str)->str:
    if not texto:
        return texto
    
    texto = EMAIL_RE.sub('[EMAIL]', texto)
    texto = PHONE_RE.sub('[PHONE]', texto)
    texto = APIKEY_RE.sub('[API_KEY]', texto)
    return texto

def decode_header_value(value):
    if not value:
        return ""
    parts=decode_header(value)
    out = []

    for bytes_or_str, charset in parts:
        try:
            if isinstance(bytes_or_str, bytes):
                out.append(bytes_or_str.decode(charset or "utf-8", errors="replace"))
            else:
                out.append(bytes_or_str)
        except Exception:
            out.append(str(bytes_or_str))
    
    return "".join(out)

def safe_decode_bytes(b: bytes, charset: str):
    if b is None:
        return
    try:
        if charset:
            return b.decode(charset, errors="replace")
        return b.decode("utf-8", errors="replace")
    except Exception:
        try:
            return b.decode("latin-1", errors="replace")
        except Exception:
            return b.decode("utf-8", errors="ignore")


def get_part_charset(part):
    cs = part.get_content_charset()
    if isinstance(cs, str):
        return cs.lower()
    return None

def extract_text_from_html(path: Path):
    html = path.read_text(encoding='utf-8', errors='ignore')
    return BeautifulSoup(html, "html.parser").get_text("\n")

def extract_text_from_docx(path: Path):
    doc = docx.Document(str(path))
    return "\n".join(p.text for p in doc.paragraphs)

def handle_generic_file(path: Path, outdir: Path, limit=None, counter=[0]):
    ext = path.suffix.lower()

    try:
        if ext=='.html' or ext=='.htm':
            text = extract_text_from_html(path)
            source='html'
        elif ext=='.pdf':
            text = extract_pdf_text(str(path)) or ""
            source = 'pdf'
        elif ext =='.docx':
            text = extract_text_from_docx(path) or ""
            source = 'docx'
        elif ext == '.json':
            try:
                raw = json.loads(path.read_text(encoding='utf-8', errors='ignore'))
                def gather_strings(o):
                    out=[]
                    if isinstance(o, str):
                        out.append(o)
                    elif isinstance(o, dict):
                        for v in o.values():
                            out.extend(gather_strings(v))
                    elif isinstance(o, list):
                        for v in o:
                            out.extend(gather_strings(v))
                    return out
                strings = gather_strings(raw)
                text="\n".join(strings[:200])
            except Exception:
                text = path.read_text(encoding='utf-8', errors='ignore')
            source='json'
        elif ext in ['.txt', '.text', '.md']:
            text = path.read_text(encoding='utf-8', errors='ignore')
            source = 'text'
        else:
            return None
        text = remove_pii(text)
        doc = {
            "id": f"{source}-{counter[0]}",
            "source": source,
            "path" : str(path),
            "text": text,
            "meta":{}
        }
        save_doc(outdir, doc)
        counter[0]+=1
        return doc["id"]
    except Exception as e:
        logging.exception(f"Failed to handle {path}: {e}")
        return None


def extraer_texto_email(msj):
    text_parts = []
    html_parts = []
    try:
        if msj.is_multipart(): #revisa si el correo tiene algo mas que solo un tipo de contenido
            for part in msj.walk():#iterates por cada parte del correo
                if part.is_multipart():#skipea containers
                    continue

                tipo_contenido = part.get_content_type() # text/plain text/html image/png
                disp = str(part.get('Content-Disposition') or "").lower()
                nombre_archivo = part.get_filename()#regresa el nombre del tipo de archivo de la seccion del correo
                if 'attachment' in disp or nombre_archivo:
                    continue#se skipean todos los attachments
                
                payload = part.get_payload(decode=True)#bytes o None
                if payload is None:
                    continue

                charset = get_part_charset(part)
                decoded_text = safe_decode_bytes(payload, charset)

                if tipo_contenido == "text/plain":
                    text_parts.append(decoded_text)
                elif tipo_contenido == "text/html":
                    html_parts.append(decoded_text)
                else:
                    continue#otro tipo como aplicaciones/json
        
        else:
            payload = msj.get_payload(decode = True)
            if payload:
                charset = msj.get_content_charset()
                tipo_contenido = msj.get_content_type()
                decoded = safe_decode_bytes(payload, charset)
                if tipo_contenido == "text/plain":
                    text_parts.append(decoded)
                elif tipo_contenido=="text/html":
                    html_parts.append(decoded)
                else:
                    text_parts.append(decoded)
    except Exception:
        return ""

    if text_parts:
        return "\n\n".join(tp.strip() for tp in text_parts if tp and tp.strip())
    

    if html_parts:
        combined_html = "\n\n".join(html_parts)
        soup = BeautifulSoup(combined_html, "html.parser")

        text = soup.get_text("\n")
        return text.strip()
    
    return ""


def parse_mbox(mbox_ubi, outdir, limit=None):
    count = 0
    mbox = mailbox.mbox(str(mbox_ubi))

    try:
        msjs_totales = len(mbox) if hasattr(mbox, "__len__") else None

        for i,msj in enumerate(mbox):
            if limit is not None and count >= limit: 
                break

            try:
                texto = extraer_texto_email(msj)

                if not texto or not texto.strip():
                    continue#msjs vacios

                subject = decode_header_value(msj.get('subject', ''))
                date = decode_header_value(msj.get('date', ''))
                sender = decode_header_value(msj.get('from', ''))

                meta = {"subject": subject, "date": date, "from": sender}

                texto = remove_pii(texto)


                doc_id = f"gmail-{mbox_ubi.stem}-{i}"

                doc = {
                    "id": doc_id,
                    "source": "gmail",
                    "path": str(mbox_ubi),
                    "timestamp": date,
                    "text": texto,
                    "meta": meta
                }

                save_doc(outdir, doc)

                count+=1

            except Exception as e:
                logging.exception(f"Failed to parse message index {i} in {mbox_ubi}: {e}")
    finally:
        try:
            mbox.close()
        except Exception:
            pass
    
    logging.info(f"Saved {count} messages from {mbox_ubi.name}")




def iterar_folder(input_path, outdir, limit_per_type=None):
    outdir.mkdir(parents=True, exist_ok=True)#se asegura que exista el folder

    logging.info(f"iterar_folder input_path={input_path} exists={input_path.exists()} is_dir={input_path.is_dir()}")

    for mbox_ubi in input_path.rglob("*.mbox"):#se crea un generator con archivos que terminan en mbox
        parse_mbox(mbox_ubi, outdir, limit=limit_per_type)

    files = list(input_path.rglob("*"))
    total = len(files)
    logging.info(f"Scanning {total} files for supported extensions (html, json, pdf, docx, txt)")

    for path in tqdm(files):
        if path.is_dir():
            continue
        if path.suffix.lower() in ['.html', '.htm', '.pdf', '.docx', '.json', '.txt', '.md']:
            handle_generic_file(path, outdir, limit=5)








if __name__=="__main__":

    # parser=argparse.ArgumentParser()
    # parser.add_argument("--input", required=True)
    # parser.add_argument("--output", default="../output")
    # parser.add_argument("--limit", type=int, default=None)
    # args=parser.parse_args()

    inpath= Path("/Users/famsz/Documents/copyd/test/Drive/Drive").expanduser().resolve()
    outdir = Path("/Users/famsz/Documents/copyd/corpus").expanduser().resolve()
    # iterar_folder(inpath, outdir, limit_per_type=args.limit)

    iterar_folder(inpath, outdir, limit_per_type=5)
    print("Parsing terminado!")