
import express from "express";
import cookieSession from "cookie-session";
import bcrypt from "bcryptjs";
import Database from "better-sqlite3";
import multer from "multer";
import fs from "fs";
import path from "path";
import crypto from "crypto";
import { Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType } from "docx";
import PptxGenJS from "pptxgenjs";
import { fileURLToPath } from "url";

const __dirname=path.dirname(fileURLToPath(import.meta.url));
const PORT=Number(process.env.PORT||3000);
const DATA_DIR=process.env.DATA_DIR||path.join(__dirname,"data");
const UPLOAD_DIR=path.join(DATA_DIR,"uploads");
const DB_PATH=path.join(DATA_DIR,"worki.db");
const MAX_UPLOAD_MB=Number(process.env.MAX_UPLOAD_MB||50);
fs.mkdirSync(DATA_DIR,{recursive:true}); fs.mkdirSync(UPLOAD_DIR,{recursive:true});
const db=new Database(DB_PATH);
db.pragma("journal_mode = WAL"); db.pragma("foreign_keys = ON");

function migrate(){
 db.exec(`
 CREATE TABLE IF NOT EXISTS clients(
  id TEXT PRIMARY KEY,name TEXT NOT NULL,biz_no TEXT,industry TEXT,contact TEXT,position TEXT,email TEXT,phone TEXT,address TEXT,website TEXT,
  plan TEXT DEFAULT 'BASIC',status TEXT DEFAULT 'Trial',monthly_fee INTEGER DEFAULT 0,monthly_ai_limit INTEGER DEFAULT 100,
  start_date TEXT,renewal_date TEXT,summary TEXT,notes TEXT,vector_store_id TEXT,created_at TEXT DEFAULT CURRENT_TIMESTAMP,updated_at TEXT DEFAULT CURRENT_TIMESTAMP
 );
 CREATE TABLE IF NOT EXISTS users(
  id TEXT PRIMARY KEY,email TEXT UNIQUE NOT NULL,password_hash TEXT NOT NULL,name TEXT,role TEXT NOT NULL CHECK(role IN ('admin','customer')),
  client_id TEXT REFERENCES clients(id) ON DELETE CASCADE,phone TEXT,active INTEGER DEFAULT 1,notes TEXT,created_at TEXT DEFAULT CURRENT_TIMESTAMP
 );
 CREATE TABLE IF NOT EXISTS documents(
  id TEXT PRIMARY KEY,client_id TEXT NOT NULL REFERENCES clients(id) ON DELETE CASCADE,name TEXT NOT NULL,category TEXT,status TEXT DEFAULT 'Uploaded',
  file_name TEXT,file_path TEXT,file_type TEXT,file_size INTEGER,summary TEXT,openai_file_id TEXT,created_at TEXT DEFAULT CURRENT_TIMESTAMP,updated_at TEXT DEFAULT CURRENT_TIMESTAMP
 );
 CREATE TABLE IF NOT EXISTS leads(
  id TEXT PRIMARY KEY,client_id TEXT NOT NULL REFERENCES clients(id) ON DELETE CASCADE,company TEXT NOT NULL,contact TEXT,phone TEXT,email TEXT,stage TEXT DEFAULT '신규',
  source TEXT,need TEXT,amount INTEGER DEFAULT 0,prob INTEGER DEFAULT 0,next_action TEXT,next_date TEXT,memo TEXT,created_at TEXT DEFAULT CURRENT_TIMESTAMP,updated_at TEXT DEFAULT CURRENT_TIMESTAMP
 );
 CREATE TABLE IF NOT EXISTS opportunities(
  id TEXT PRIMARY KEY,client_id TEXT NOT NULL REFERENCES clients(id) ON DELETE CASCADE,type TEXT,title TEXT NOT NULL,issuer TEXT,notice_date TEXT,deadline TEXT,budget TEXT,url TEXT,
  fit TEXT DEFAULT 'Medium',status TEXT DEFAULT '신규',manager TEXT,next_action TEXT,eligibility TEXT,deliverables TEXT,memo TEXT,
  attachment_name TEXT,attachment_path TEXT,created_at TEXT DEFAULT CURRENT_TIMESTAMP,updated_at TEXT DEFAULT CURRENT_TIMESTAMP
 );
 CREATE TABLE IF NOT EXISTS tickets(
  id TEXT PRIMARY KEY,client_id TEXT NOT NULL REFERENCES clients(id) ON DELETE CASCADE,sender TEXT,email TEXT,category TEXT,priority TEXT,status TEXT,message TEXT,draft TEXT,owner TEXT,
  created_at TEXT DEFAULT CURRENT_TIMESTAMP,updated_at TEXT DEFAULT CURRENT_TIMESTAMP
 );
 CREATE TABLE IF NOT EXISTS reports(
  id TEXT PRIMARY KEY,client_id TEXT NOT NULL REFERENCES clients(id) ON DELETE CASCADE,period TEXT,kpi TEXT,risks TEXT,actions TEXT,status TEXT,memo TEXT,
  created_at TEXT DEFAULT CURRENT_TIMESTAMP,updated_at TEXT DEFAULT CURRENT_TIMESTAMP
 );
 CREATE TABLE IF NOT EXISTS ai_usage(
  id TEXT PRIMARY KEY,client_id TEXT NOT NULL REFERENCES clients(id) ON DELETE CASCADE,user_id TEXT REFERENCES users(id) ON DELETE SET NULL,role TEXT,request TEXT,result TEXT,
  input_tokens INTEGER DEFAULT 0,output_tokens INTEGER DEFAULT 0,created_at TEXT DEFAULT CURRENT_TIMESTAMP
 );
 CREATE TABLE IF NOT EXISTS automations(
  id TEXT PRIMARY KEY,name TEXT NOT NULL,trigger_name TEXT,condition_text TEXT,action_text TEXT,status TEXT DEFAULT 'ON',owner TEXT,notes TEXT,
  created_at TEXT DEFAULT CURRENT_TIMESTAMP,updated_at TEXT DEFAULT CURRENT_TIMESTAMP
 );
 CREATE TABLE IF NOT EXISTS audit_logs(
  id TEXT PRIMARY KEY,user_id TEXT,action TEXT,entity TEXT,entity_id TEXT,detail TEXT,created_at TEXT DEFAULT CURRENT_TIMESTAMP
 );
 CREATE TABLE IF NOT EXISTS ai_results(
  id TEXT PRIMARY KEY,
  client_id TEXT NOT NULL REFERENCES clients(id) ON DELETE CASCADE,
  user_id TEXT REFERENCES users(id) ON DELETE SET NULL,
  role TEXT NOT NULL,
  deliverable_type TEXT,
  title TEXT,
  request_text TEXT,
  input_json TEXT,
  result_json TEXT,
  result_text TEXT,
  model TEXT,
  created_at TEXT DEFAULT CURRENT_TIMESTAMP,
  updated_at TEXT DEFAULT CURRENT_TIMESTAMP
 );
 `);
}
migrate();
const uid=(p)=>`${p}_${crypto.randomBytes(6).toString("hex")}`;
function seed(){
 const c=db.prepare("SELECT COUNT(*) n FROM clients").get().n;
 if(!c){
  db.prepare(`INSERT INTO clients(id,name,biz_no,industry,contact,position,email,phone,address,website,plan,status,monthly_fee,monthly_ai_limit,start_date,renewal_date,summary,notes)
  VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)`).run("C1","ABC건설","123-45-67890","건설·엔지니어링","홍길동","대표","ceo@abcbuilder.demo","010-1111-2222","서울특별시","","PRO","Active",790000,300,"2026-10-01","2026-11-01","공공입찰·기술제안서·거래처 영업 중심 기업","핵심 고객");
  db.prepare(`INSERT INTO clients(id,name,industry,contact,position,email,phone,plan,status,monthly_fee,monthly_ai_limit,start_date,renewal_date,summary)
  VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)`).run("C2","미래산업","제조","김영업","팀장","sales@mirai.demo","010-3333-4444","BASIC","Trial",390000,100,"2026-10-03","2026-11-03","산업용 부품 제조 및 B2B 영업");
 }
 const u=db.prepare("SELECT COUNT(*) n FROM users").get().n;
 if(!u){
  const adminEmail=process.env.ADMIN_EMAIL||"admin@worki.local";
  const adminPw=process.env.ADMIN_PASSWORD||"ChangeMe123!";
  db.prepare("INSERT INTO users(id,email,password_hash,name,role,active,notes) VALUES(?,?,?,?,?,?,?)").run("U_ADMIN",adminEmail,bcrypt.hashSync(adminPw,10),"WORKI 관리자","admin",1,"초기 관리자");
  db.prepare("INSERT INTO users(id,email,password_hash,name,role,client_id,active) VALUES(?,?,?,?,?,?,?)").run("U_C1","ceo@abcbuilder.demo",bcrypt.hashSync("1234",10),"홍길동 대표","customer","C1",1);
  db.prepare("INSERT INTO users(id,email,password_hash,name,role,client_id,active) VALUES(?,?,?,?,?,?,?)").run("U_C2","sales@mirai.demo",bcrypt.hashSync("1234",10),"김영업 팀장","customer","C2",1);
 }
 if(db.prepare("SELECT COUNT(*) n FROM leads").get().n===0){
  db.prepare(`INSERT INTO leads(id,client_id,company,contact,email,stage,source,need,amount,prob,next_action,next_date,memo) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)`)
    .run("L1","C1","대한엔지니어링","이이사","lee@example.com","제안","소개","입찰·제안서 자동화",4900000,80,"제안서 피드백","2026-10-12","의사결정 빠름");
 }
 if(db.prepare("SELECT COUNT(*) n FROM opportunities").get().n===0){
  db.prepare(`INSERT INTO opportunities(id,client_id,type,title,issuer,notice_date,deadline,budget,fit,status,manager,next_action,eligibility,deliverables,memo)
  VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)`).run("O1","C1","정부지원사업","스마트건설 기술지원 사업","중소벤처기업부","2026-10-01","2026-10-31","최대 5천만원","High","검토중","홍길동","자격요건 확인","중소기업, 관련업종","신청서, 사업계획서","샘플 데이터");
 }
}
seed();

const app=express();
app.use(express.json({limit:"2mb"}));
app.use(express.urlencoded({extended:true}));
app.use(cookieSession({name:"worki_session",keys:[process.env.SESSION_SECRET||"dev-secret-change-me"],httpOnly:true,sameSite:"lax",secure:process.env.NODE_ENV==="production",maxAge:1000*60*60*12}));
app.use(express.static(path.join(__dirname,"public")));
const upload=multer({dest:UPLOAD_DIR,limits:{fileSize:MAX_UPLOAD_MB*1024*1024}});

function auth(req,res,next){if(!req.session?.user)return res.status(401).json({error:"로그인이 필요합니다."});const u=db.prepare("SELECT id,email,name,role,client_id,active FROM users WHERE id=?").get(req.session.user.id);if(!u||!u.active)return res.status(401).json({error:"비활성 계정입니다."});req.user=u;next()}
function admin(req,res,next){if(req.user.role!=="admin")return res.status(403).json({error:"관리자 권한이 필요합니다."});next()}
function tenantClause(req,field="client_id"){return req.user.role==="admin"?{sql:"1=1",params:[]}:{sql:`${field}=?`,params:[req.user.client_id]}}
function canClient(req,cid){return req.user.role==="admin"||req.user.client_id===cid}
function log(req,action,entity,entityId,detail=""){try{db.prepare("INSERT INTO audit_logs(id,user_id,action,entity,entity_id,detail) VALUES(?,?,?,?,?,?)").run(uid("AUD"),req.user?.id||null,action,entity,entityId||null,detail)}catch{}}
function cleanFile(p){if(p&&fs.existsSync(p))try{fs.unlinkSync(p)}catch{}}

app.post("/api/login",(req,res)=>{
 const u=db.prepare("SELECT * FROM users WHERE email=?").get(req.body.email);
 if(!u||!u.active||!bcrypt.compareSync(req.body.password||"",u.password_hash)) return res.status(401).json({error:"이메일 또는 비밀번호가 올바르지 않습니다."});
 req.session.user={id:u.id};res.json({ok:true,user:{id:u.id,email:u.email,name:u.name,role:u.role,client_id:u.client_id}});
});
app.post("/api/logout",(req,res)=>{req.session=null;res.json({ok:true})});
app.get("/api/me",auth,(req,res)=>res.json(req.user));

app.get("/api/dashboard",auth,(req,res)=>{
 const c=tenantClause(req,"id"); const tc=tenantClause(req);
 const clients=req.user.role==="admin"?db.prepare("SELECT COUNT(*) n FROM clients").get().n:1;
 const docs=db.prepare(`SELECT COUNT(*) n FROM documents WHERE ${tc.sql}`).get(...tc.params).n;
 const leads=db.prepare(`SELECT COUNT(*) n,COALESCE(SUM(amount),0) amount FROM leads WHERE ${tc.sql}`).get(...tc.params);
 const opps=db.prepare(`SELECT COUNT(*) n FROM opportunities WHERE ${tc.sql}`).get(...tc.params).n;
 const usage=db.prepare(`SELECT COUNT(*) n FROM ai_usage WHERE ${tc.sql} AND date(created_at)>=date('now','start of month')`).get(...tc.params).n;
 res.json({clients,docs,leads:leads.n,pipeline:leads.amount,opps,usage});
});

app.get("/api/clients",auth,(req,res)=>{
 if(req.user.role==="admin") return res.json(db.prepare("SELECT * FROM clients ORDER BY created_at DESC").all());
 res.json(db.prepare("SELECT * FROM clients WHERE id=?").all(req.user.client_id));
});
app.post("/api/clients",auth,admin,(req,res)=>{const id=uid("C");const b=req.body;db.prepare(`INSERT INTO clients(id,name,biz_no,industry,contact,position,email,phone,address,website,plan,status,monthly_fee,monthly_ai_limit,start_date,renewal_date,summary,notes) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)`).run(id,b.name,b.biz_no,b.industry,b.contact,b.position,b.email,b.phone,b.address,b.website,b.plan||"BASIC",b.status||"Trial",Number(b.monthly_fee||0),Number(b.monthly_ai_limit||100),b.start_date,b.renewal_date,b.summary,b.notes);log(req,"create","client",id,b.name);res.json({id})});
app.put("/api/clients/:id",auth,admin,(req,res)=>{const b=req.body;db.prepare(`UPDATE clients SET name=?,biz_no=?,industry=?,contact=?,position=?,email=?,phone=?,address=?,website=?,plan=?,status=?,monthly_fee=?,monthly_ai_limit=?,start_date=?,renewal_date=?,summary=?,notes=?,updated_at=CURRENT_TIMESTAMP WHERE id=?`).run(b.name,b.biz_no,b.industry,b.contact,b.position,b.email,b.phone,b.address,b.website,b.plan,b.status,Number(b.monthly_fee||0),Number(b.monthly_ai_limit||0),b.start_date,b.renewal_date,b.summary,b.notes,req.params.id);log(req,"update","client",req.params.id,b.name);res.json({ok:true})});
app.delete("/api/clients/:id",auth,admin,(req,res)=>{const docs=db.prepare("SELECT file_path FROM documents WHERE client_id=?").all(req.params.id);docs.forEach(x=>cleanFile(x.file_path));const opps=db.prepare("SELECT attachment_path FROM opportunities WHERE client_id=?").all(req.params.id);opps.forEach(x=>cleanFile(x.attachment_path));db.prepare("DELETE FROM clients WHERE id=?").run(req.params.id);log(req,"delete","client",req.params.id);res.json({ok:true})});

app.get("/api/users",auth,admin,(req,res)=>res.json(db.prepare("SELECT id,email,name,role,client_id,phone,active,notes,created_at FROM users ORDER BY created_at DESC").all()));
app.post("/api/users",auth,admin,(req,res)=>{const b=req.body,id=uid("U");db.prepare("INSERT INTO users(id,email,password_hash,name,role,client_id,phone,active,notes) VALUES(?,?,?,?,?,?,?,?,?)").run(id,b.email,bcrypt.hashSync(b.password||"1234",10),b.name,b.role||"customer",b.client_id||null,b.phone,b.active===false?0:1,b.notes);log(req,"create","user",id,b.email);res.json({id})});
app.put("/api/users/:id",auth,admin,(req,res)=>{const b=req.body;if(b.password)db.prepare("UPDATE users SET password_hash=? WHERE id=?").run(bcrypt.hashSync(b.password,10),req.params.id);db.prepare("UPDATE users SET email=?,name=?,role=?,client_id=?,phone=?,active=?,notes=? WHERE id=?").run(b.email,b.name,b.role,b.client_id||null,b.phone,b.active?1:0,b.notes,req.params.id);log(req,"update","user",req.params.id,b.email);res.json({ok:true})});
app.delete("/api/users/:id",auth,admin,(req,res)=>{if(req.params.id===req.user.id)return res.status(400).json({error:"현재 로그인 계정은 삭제할 수 없습니다."});db.prepare("DELETE FROM users WHERE id=?").run(req.params.id);log(req,"delete","user",req.params.id);res.json({ok:true})});

app.get("/api/documents",auth,(req,res)=>{const c=tenantClause(req);res.json(db.prepare(`SELECT * FROM documents WHERE ${c.sql} ORDER BY created_at DESC`).all(...c.params))});
app.post("/api/documents",auth,upload.single("file"),async(req,res)=>{
 const cid=req.body.client_id||req.user.client_id;if(!canClient(req,cid)){cleanFile(req.file?.path);return res.status(403).json({error:"권한 없음"})}
 const id=uid("D"),f=req.file;db.prepare(`INSERT INTO documents(id,client_id,name,category,status,file_name,file_path,file_type,file_size,summary) VALUES(?,?,?,?,?,?,?,?,?,?)`).run(id,cid,req.body.name||f?.originalname||"자료",req.body.category||"기타","Uploaded",f?.originalname||"",f?.path||"",f?.mimetype||"",f?.size||0,req.body.summary||"");
 log(req,"create","document",id,f?.originalname||"");res.json({id});
});
app.put("/api/documents/:id",auth,(req,res)=>{const d=db.prepare("SELECT * FROM documents WHERE id=?").get(req.params.id);if(!d||!canClient(req,d.client_id))return res.status(404).json({error:"자료 없음"});const b=req.body;db.prepare("UPDATE documents SET name=?,category=?,status=?,summary=?,updated_at=CURRENT_TIMESTAMP WHERE id=?").run(b.name,b.category,b.status,b.summary,req.params.id);log(req,"update","document",req.params.id);res.json({ok:true})});
app.get("/api/documents/:id/download",auth,(req,res)=>{const d=db.prepare("SELECT * FROM documents WHERE id=?").get(req.params.id);if(!d||!canClient(req,d.client_id)||!d.file_path||!fs.existsSync(d.file_path))return res.status(404).send("파일 없음");res.download(d.file_path,d.file_name)});
app.delete("/api/documents/:id",auth,(req,res)=>{const d=db.prepare("SELECT * FROM documents WHERE id=?").get(req.params.id);if(!d||!canClient(req,d.client_id))return res.status(404).json({error:"자료 없음"});cleanFile(d.file_path);db.prepare("DELETE FROM documents WHERE id=?").run(req.params.id);log(req,"delete","document",req.params.id);res.json({ok:true})});

function crudRoutes(base,table,fields,{adminOnly=false}={}){
 app.get(`/api/${base}`,auth,(req,res)=>{const c=tenantClause(req);res.json(db.prepare(`SELECT * FROM ${table} WHERE ${c.sql} ORDER BY created_at DESC`).all(...c.params))});
 app.post(`/api/${base}`,auth,(req,res)=>{if(adminOnly&&req.user.role!=="admin")return res.status(403).json({error:"관리자 권한 필요"});const cid=req.body.client_id||req.user.client_id;if("client_id" in req.body||table!=="automations"){if(table!=="automations"&&!canClient(req,cid))return res.status(403).json({error:"권한 없음"})};const id=uid(base.slice(0,3).toUpperCase());const cols=fields.filter(x=>x!=="id");const vals=cols.map(x=>x==="client_id"?cid:req.body[x]??null);db.prepare(`INSERT INTO ${table}(id,${cols.join(",")}) VALUES(?,${cols.map(()=>"?").join(",")})`).run(id,...vals);log(req,"create",table,id);res.json({id})});
 app.put(`/api/${base}/:id`,auth,(req,res)=>{const row=db.prepare(`SELECT * FROM ${table} WHERE id=?`).get(req.params.id);if(!row)return res.status(404).json({error:"없음"});if(table!=="automations"&&!canClient(req,row.client_id))return res.status(403).json({error:"권한 없음"});if(adminOnly&&req.user.role!=="admin")return res.status(403).json({error:"관리자 권한 필요"});const cols=fields.filter(x=>!["id","client_id"].includes(x));const vals=cols.map(x=>req.body[x]??row[x]??null);db.prepare(`UPDATE ${table} SET ${cols.map(x=>`${x}=?`).join(",")},updated_at=CURRENT_TIMESTAMP WHERE id=?`).run(...vals,req.params.id);log(req,"update",table,req.params.id);res.json({ok:true})});
 app.delete(`/api/${base}/:id`,auth,(req,res)=>{const row=db.prepare(`SELECT * FROM ${table} WHERE id=?`).get(req.params.id);if(!row)return res.status(404).json({error:"없음"});if(table!=="automations"&&!canClient(req,row.client_id))return res.status(403).json({error:"권한 없음"});if(adminOnly&&req.user.role!=="admin")return res.status(403).json({error:"관리자 권한 필요"});db.prepare(`DELETE FROM ${table} WHERE id=?`).run(req.params.id);log(req,"delete",table,req.params.id);res.json({ok:true})});
}
crudRoutes("leads","leads",["id","client_id","company","contact","phone","email","stage","source","need","amount","prob","next_action","next_date","memo"]);
crudRoutes("tickets","tickets",["id","client_id","sender","email","category","priority","status","message","draft","owner"]);
crudRoutes("reports","reports",["id","client_id","period","kpi","risks","actions","status","memo"]);
crudRoutes("automations","automations",["id","name","trigger_name","condition_text","action_text","status","owner","notes"],{adminOnly:true});

app.get("/api/opportunities",auth,(req,res)=>{const c=tenantClause(req);res.json(db.prepare(`SELECT * FROM opportunities WHERE ${c.sql} ORDER BY COALESCE(deadline,'9999-12-31'),created_at DESC`).all(...c.params))});
app.post("/api/opportunities",auth,upload.single("attachment"),(req,res)=>{const cid=req.body.client_id||req.user.client_id;if(!canClient(req,cid)){cleanFile(req.file?.path);return res.status(403).json({error:"권한 없음"})}const id=uid("OPP"),f=req.file,b=req.body;db.prepare(`INSERT INTO opportunities(id,client_id,type,title,issuer,notice_date,deadline,budget,url,fit,status,manager,next_action,eligibility,deliverables,memo,attachment_name,attachment_path) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)`).run(id,cid,b.type,b.title,b.issuer,b.notice_date,b.deadline,b.budget,b.url,b.fit||"Medium",b.status||"신규",b.manager,b.next_action,b.eligibility,b.deliverables,b.memo,f?.originalname||"",f?.path||"");log(req,"create","opportunity",id,b.title);res.json({id})});
app.put("/api/opportunities/:id",auth,(req,res)=>{const o=db.prepare("SELECT * FROM opportunities WHERE id=?").get(req.params.id);if(!o||!canClient(req,o.client_id))return res.status(404).json({error:"없음"});const b=req.body;db.prepare(`UPDATE opportunities SET type=?,title=?,issuer=?,notice_date=?,deadline=?,budget=?,url=?,fit=?,status=?,manager=?,next_action=?,eligibility=?,deliverables=?,memo=?,updated_at=CURRENT_TIMESTAMP WHERE id=?`).run(b.type,b.title,b.issuer,b.notice_date,b.deadline,b.budget,b.url,b.fit,b.status,b.manager,b.next_action,b.eligibility,b.deliverables,b.memo,req.params.id);log(req,"update","opportunity",req.params.id);res.json({ok:true})});
app.get("/api/opportunities/:id/download",auth,(req,res)=>{const o=db.prepare("SELECT * FROM opportunities WHERE id=?").get(req.params.id);if(!o||!canClient(req,o.client_id)||!o.attachment_path||!fs.existsSync(o.attachment_path))return res.status(404).send("첨부 없음");res.download(o.attachment_path,o.attachment_name)});
app.delete("/api/opportunities/:id",auth,(req,res)=>{const o=db.prepare("SELECT * FROM opportunities WHERE id=?").get(req.params.id);if(!o||!canClient(req,o.client_id))return res.status(404).json({error:"없음"});cleanFile(o.attachment_path);db.prepare("DELETE FROM opportunities WHERE id=?").run(req.params.id);log(req,"delete","opportunity",req.params.id);res.json({ok:true})});

async function openai(pathname,opts={}){
 const key=process.env.OPENAI_API_KEY;if(!key)throw new Error("OPENAI_API_KEY가 설정되지 않았습니다.");
 const r=await fetch("https://api.openai.com/v1"+pathname,{...opts,headers:{Authorization:`Bearer ${key}`,...(opts.headers||{})}});
 const d=await r.json();if(!r.ok)throw new Error(d?.error?.message||`OpenAI ${r.status}`);return d;
}
app.post("/api/documents/:id/ingest",auth,async(req,res)=>{
 try{
  const d=db.prepare("SELECT d.*,c.vector_store_id,c.name client_name FROM documents d JOIN clients c ON c.id=d.client_id WHERE d.id=?").get(req.params.id);
  if(!d||!canClient(req,d.client_id)||!d.file_path||!fs.existsSync(d.file_path))return res.status(404).json({error:"파일 없음"});
  let vs=d.vector_store_id;
  if(!vs){const v=await openai("/vector_stores",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({name:`WORKI - ${d.client_name}`})});vs=v.id;db.prepare("UPDATE clients SET vector_store_id=? WHERE id=?").run(vs,d.client_id)}
  const buf=fs.readFileSync(d.file_path),blob=new Blob([buf],{type:d.file_type||"application/octet-stream"}),form=new FormData();form.append("purpose","assistants");form.append("file",blob,d.file_name);
  const f=await openai("/files",{method:"POST",body:form});
  await openai(`/vector_stores/${vs}/files`,{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({file_id:f.id,attributes:{client_id:d.client_id,document_id:d.id}})});
  db.prepare("UPDATE documents SET status='Ready',openai_file_id=?,updated_at=CURRENT_TIMESTAMP WHERE id=?").run(f.id,d.id);log(req,"ingest","document",d.id);res.json({ok:true,vector_store_id:vs,file_id:f.id});
 }catch(e){db.prepare("UPDATE documents SET status='Failed',updated_at=CURRENT_TIMESTAMP WHERE id=?").run(req.params.id);res.status(500).json({error:e.message})}
});

function safeJsonParse(text){
  if(!text)return null;
  try{return JSON.parse(text)}catch{}
  const m=text.match(/\{[\s\S]*\}/);
  if(m){try{return JSON.parse(m[0])}catch{}}
  return null;
}
function roleSpec(role){
 const common=`반드시 한국어로 작성한다. 회사자료가 검색되면 그 자료를 우선 근거로 사용한다. 자료에 없는 숫자, 실적, 고객명, 인증, 계약조건은 만들어내지 않는다. 사용자가 바로 업무에 사용할 수 있을 정도로 구체적으로 작성한다. 출력은 설명문 없이 유효한 JSON 객체 하나만 반환한다.`;
 const specs={
 "AI 영업사원":`${common}
JSON 구조:
{"title":"","executive_summary":"","sections":[{"heading":"고객 분석","body":"","bullets":[]},{"heading":"가치제안","body":"","bullets":[]},{"heading":"영업 전략","body":"","bullets":[]}],"email":{"subject":"","body":""},"followup_email":{"subject":"","body":""},"call_script":{"opening":"","questions":[],"objection_handling":[],"closing":""},"action_plan":[{"step":"","owner":"","due":"","note":""}],"checklist":[],"slides":[]}`,
 "AI 문서비서":`${common}
JSON 구조:
{"title":"","executive_summary":"","sections":[{"heading":"","body":"","bullets":[]}],"email":{"subject":"","body":""},"action_plan":[{"step":"","owner":"","due":"","note":""}],"checklist":[],"slides":[]}`,
 "AI 정보조사원":`${common}
JSON 구조:
{"title":"","executive_summary":"","fit":"High|Medium|Low","sections":[{"heading":"핵심 내용","body":"","bullets":[]},{"heading":"자격/조건","body":"","bullets":[]},{"heading":"리스크","body":"","bullets":[]},{"heading":"추천 대응","body":"","bullets":[]}],"checklist":[],"action_plan":[{"step":"","owner":"","due":"","note":""}],"table":[{"항목":"","내용":"","확인상태":""}],"slides":[]}`,
 "AI 제안서 기획자":`${common}
JSON 구조:
{"title":"","executive_summary":"","sections":[{"heading":"고객문제","body":"","bullets":[]},{"heading":"제안전략","body":"","bullets":[]},{"heading":"솔루션","body":"","bullets":[]},{"heading":"수행범위","body":"","bullets":[]},{"heading":"기대효과","body":"","bullets":[]}],"checklist":[],"action_plan":[{"step":"","owner":"","due":"","note":""}],"slides":[{"title":"","subtitle":"","bullets":[],"speaker_notes":""}]}
slides는 최소 10장, 최대 15장으로 구성하고 각 슬라이드에 실제 발표자료에 넣을 수 있는 구체적 bullet을 작성한다.`
 };
 return specs[role]||specs["AI 문서비서"];
}
function resultToPlain(r){
 const out=[];
 if(r.title)out.push(r.title,"\n");
 if(r.executive_summary)out.push("요약\n"+r.executive_summary+"\n");
 for(const s of (r.sections||[])){out.push("\n"+(s.heading||"섹션")+"\n");if(s.body)out.push(s.body+"\n");for(const b of (s.bullets||[]))out.push("- "+b)}
 if(r.email?.subject||r.email?.body)out.push("\n영업/업무 이메일\n제목: "+(r.email.subject||"")+"\n"+(r.email.body||""));
 if(r.followup_email?.subject||r.followup_email?.body)out.push("\n후속 이메일\n제목: "+(r.followup_email.subject||"")+"\n"+(r.followup_email.body||""));
 if(r.call_script){out.push("\n콜 스크립트\n"+(r.call_script.opening||""));for(const q of (r.call_script.questions||[]))out.push("- 질문: "+q);for(const o of (r.call_script.objection_handling||[]))out.push("- 반론대응: "+o);if(r.call_script.closing)out.push("마무리: "+r.call_script.closing)}
 if(r.action_plan?.length){out.push("\n실행계획");r.action_plan.forEach((a,i)=>out.push(`${i+1}. ${a.step||""} | 담당 ${a.owner||"-"} | 기한 ${a.due||"-"} | ${a.note||""}`))}
 if(r.checklist?.length){out.push("\n체크리스트");r.checklist.forEach(x=>out.push("- [ ] "+(typeof x==="string"?x:JSON.stringify(x))))}
 if(r.slides?.length){out.push("\n제안서 슬라이드 구성");r.slides.forEach((s,i)=>{out.push(`\n${i+1}. ${s.title||""}`);if(s.subtitle)out.push(s.subtitle);for(const b of (s.bullets||[]))out.push("- "+b)})}
 return out.join("\n");
}
async function buildDocx(result,meta){
 const children=[
  new Paragraph({text:result.title||meta.title||"WORKI AI 결과물",heading:HeadingLevel.TITLE}),
  new Paragraph({children:[new TextRun({text:`AI 직원: ${meta.role} | 고객사: ${meta.clientName} | 생성일: ${new Date().toLocaleDateString("ko-KR")}`,italics:true,color:"666666"})]}),
 ];
 if(result.executive_summary){
  children.push(new Paragraph({text:"Executive Summary",heading:HeadingLevel.HEADING_1}));
  children.push(new Paragraph(result.executive_summary));
 }
 for(const s of (result.sections||[])){
  children.push(new Paragraph({text:s.heading||"내용",heading:HeadingLevel.HEADING_1}));
  if(s.body)children.push(new Paragraph(s.body));
  for(const b of (s.bullets||[]))children.push(new Paragraph({text:b,bullet:{level:0}}));
 }
 if(result.email?.subject||result.email?.body){
  children.push(new Paragraph({text:"이메일 완성문안",heading:HeadingLevel.HEADING_1}));
  children.push(new Paragraph({children:[new TextRun({text:"제목: ",bold:true}),new TextRun(result.email.subject||"")]}));
  children.push(new Paragraph(result.email.body||""));
 }
 if(result.followup_email?.subject||result.followup_email?.body){
  children.push(new Paragraph({text:"후속 이메일",heading:HeadingLevel.HEADING_1}));
  children.push(new Paragraph({children:[new TextRun({text:"제목: ",bold:true}),new TextRun(result.followup_email.subject||"")]}));
  children.push(new Paragraph(result.followup_email.body||""));
 }
 if(result.call_script){
  children.push(new Paragraph({text:"콜 스크립트",heading:HeadingLevel.HEADING_1}));
  if(result.call_script.opening)children.push(new Paragraph(result.call_script.opening));
  for(const q of (result.call_script.questions||[]))children.push(new Paragraph({text:"질문: "+q,bullet:{level:0}}));
  for(const o of (result.call_script.objection_handling||[]))children.push(new Paragraph({text:"반론 대응: "+o,bullet:{level:0}}));
  if(result.call_script.closing)children.push(new Paragraph("마무리: "+result.call_script.closing));
 }
 if(result.action_plan?.length){
  children.push(new Paragraph({text:"실행계획",heading:HeadingLevel.HEADING_1}));
  const rows=[new TableRow({children:["순서","업무","담당","기한","메모"].map(x=>new TableCell({children:[new Paragraph({children:[new TextRun({text:x,bold:true})]})]}))})];
  result.action_plan.forEach((a,i)=>rows.push(new TableRow({children:[String(i+1),a.step||"",a.owner||"",a.due||"",a.note||""].map(x=>new TableCell({children:[new Paragraph(String(x))]}))})));
  children.push(new Table({rows,width:{size:100,type:WidthType.PERCENTAGE}}));
 }
 if(result.checklist?.length){
  children.push(new Paragraph({text:"체크리스트",heading:HeadingLevel.HEADING_1}));
  for(const x of result.checklist)children.push(new Paragraph({text:"☐ "+(typeof x==="string"?x:JSON.stringify(x))}));
 }
 if(result.slides?.length){
  children.push(new Paragraph({text:"제안서 슬라이드 구성",heading:HeadingLevel.HEADING_1}));
  result.slides.forEach((s,i)=>{
   children.push(new Paragraph({text:`${i+1}. ${s.title||""}`,heading:HeadingLevel.HEADING_2}));
   if(s.subtitle)children.push(new Paragraph(s.subtitle));
   for(const b of (s.bullets||[]))children.push(new Paragraph({text:b,bullet:{level:0}}));
   if(s.speaker_notes)children.push(new Paragraph({children:[new TextRun({text:"발표메모: ",bold:true}),new TextRun(s.speaker_notes)]}));
  });
 }
 const doc=new Document({sections:[{properties:{},children}]});
 return await Packer.toBuffer(doc);
}
async function buildPptx(result,meta){
 const pptx=new PptxGenJS();pptx.layout="LAYOUT_WIDE";pptx.author="WORKI";pptx.subject=meta.role;pptx.title=result.title||"WORKI AI 결과물";pptx.company="WORKI";pptx.lang="ko-KR";
 pptx.theme={headFontFace:"Aptos",bodyFontFace:"Aptos",lang:"ko-KR"};
 let slide=pptx.addSlide();slide.background={color:"0B1B2D"};slide.addText(result.title||"WORKI AI 결과물",{x:.7,y:1.6,w:11.8,h:.7,fontSize:28,bold:true,color:"FFFFFF"});slide.addText(`${meta.clientName} | ${meta.role}`,{x:.7,y:2.5,w:11.8,h:.4,fontSize:15,color:"67E8F9"});slide.addText(result.executive_summary||"",{x:.7,y:3.2,w:11.5,h:1.7,fontSize:16,color:"C6D4E3",breakLine:false});
 const slides=result.slides?.length?result.slides:(result.sections||[]).map(s=>({title:s.heading,bullets:[s.body,...(s.bullets||[])].filter(Boolean)}));
 for(const s of slides){
  slide=pptx.addSlide();slide.background={color:"F7FAFC"};slide.addText(s.title||"",{x:.6,y:.45,w:12,h:.5,fontSize:24,bold:true,color:"0F2740"});if(s.subtitle)slide.addText(s.subtitle,{x:.65,y:1.1,w:11.8,h:.35,fontSize:13,color:"52677C"});
  const bullets=(s.bullets||[]).slice(0,8);slide.addText(bullets.map(x=>({text:String(x),options:{bullet:{indent:18},breakLine:true}})),{x:.8,y:1.65,w:11.2,h:5.2,fontSize:17,color:"23384E",breakLine:false,paraSpaceAfterPt:10,valign:"top"});
  slide.addText("WORKI",{x:11.7,y:7.05,w:1,h:.2,fontSize:8,color:"7990A7"});
 }
 return await pptx.write({outputType:"nodebuffer"});
}
function toCsv(result){
 const rows=[["구분","항목","내용"]];
 rows.push(["기본","제목",result.title||""]);rows.push(["기본","요약",result.executive_summary||""]);
 (result.sections||[]).forEach(s=>{rows.push(["섹션",s.heading||"",s.body||""]);(s.bullets||[]).forEach(b=>rows.push(["bullet",s.heading||"",b]))});
 (result.action_plan||[]).forEach(a=>rows.push(["실행계획",a.step||"",`${a.owner||""} | ${a.due||""} | ${a.note||""}`]));
 (result.checklist||[]).forEach(x=>rows.push(["체크리스트","",typeof x==="string"?x:JSON.stringify(x)]));
 (result.slides||[]).forEach((s,i)=>rows.push(["슬라이드",`${i+1}. ${s.title||""}`,(s.bullets||[]).join(" / ")]));
 const q=v=>`"${String(v??"").replace(/"/g,'""')}"`;
 return "\ufeff"+rows.map(r=>r.map(q).join(",")).join("\r\n");
}
app.post("/api/ai/generate",auth,async(req,res)=>{
 try{
  const cid=req.body.client_id||req.user.client_id;if(!canClient(req,cid))return res.status(403).json({error:"권한 없음"});
  const c=db.prepare("SELECT * FROM clients WHERE id=?").get(cid);if(!c)return res.status(404).json({error:"고객사 없음"});
  const role=req.body.role||"AI 문서비서";
  const input={
   goal:req.body.goal||"",audience:req.body.audience||"",background:req.body.background||"",
   requirements:req.body.requirements||"",tone:req.body.tone||"전문적",length:req.body.length||"상세",
   deliverable_type:req.body.deliverable_type||"업무문서"
  };
  const userPrompt=`다음 조건으로 완성 결과물을 작성하라.
고객사: ${c.name}
업종: ${c.industry||""}
회사설명: ${c.summary||""}
결과물 종류: ${input.deliverable_type}
목적: ${input.goal}
대상/수신자: ${input.audience}
배경/상황: ${input.background}
핵심 요구사항: ${input.requirements}
톤: ${input.tone}
분량: ${input.length}`;
  const body={model:process.env.OPENAI_MODEL||"gpt-6-luna",input:[{role:"system",content:roleSpec(role)},{role:"user",content:userPrompt}],max_output_tokens:8000};
  if(c.vector_store_id)body.tools=[{type:"file_search",vector_store_ids:[c.vector_store_id],max_num_results:10}];
  const data=await openai("/responses",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify(body)});
  const raw=data.output_text||(data.output||[]).flatMap(x=>x.content||[]).filter(x=>x.type==="output_text").map(x=>x.text).join("\n");
  let result=safeJsonParse(raw);
  if(!result)result={title:`${role} 결과물`,executive_summary:"",sections:[{heading:"AI 결과",body:raw,bullets:[]}],checklist:[],action_plan:[],slides:[]};
  const id=uid("RES"),title=result.title||`${role} 결과물`,plain=resultToPlain(result);
  db.prepare(`INSERT INTO ai_results(id,client_id,user_id,role,deliverable_type,title,request_text,input_json,result_json,result_text,model)
   VALUES(?,?,?,?,?,?,?,?,?,?,?)`).run(id,cid,req.user.id,role,input.deliverable_type,title,userPrompt,JSON.stringify(input),JSON.stringify(result),plain,data.model||process.env.OPENAI_MODEL||"gpt-6-luna");
  db.prepare("INSERT INTO ai_usage(id,client_id,user_id,role,request,result,input_tokens,output_tokens) VALUES(?,?,?,?,?,?,?,?)").run(uid("AI"),cid,req.user.id,role,userPrompt,plain,data.usage?.input_tokens||0,data.usage?.output_tokens||0);
  res.json({id,title,result,plain,usage:data.usage||{}});
 }catch(e){res.status(500).json({error:e.message})}
});
app.get("/api/results",auth,(req,res)=>{
 const c=tenantClause(req);
 res.json(db.prepare(`SELECT id,client_id,user_id,role,deliverable_type,title,model,created_at,updated_at FROM ai_results WHERE ${c.sql} ORDER BY created_at DESC LIMIT 200`).all(...c.params));
});
app.get("/api/results/:id",auth,(req,res)=>{
 const r=db.prepare("SELECT * FROM ai_results WHERE id=?").get(req.params.id);
 if(!r||!canClient(req,r.client_id))return res.status(404).json({error:"결과물 없음"});
 r.result=JSON.parse(r.result_json||"{}");r.input=JSON.parse(r.input_json||"{}");delete r.result_json;delete r.input_json;res.json(r);
});
app.delete("/api/results/:id",auth,(req,res)=>{
 const r=db.prepare("SELECT * FROM ai_results WHERE id=?").get(req.params.id);
 if(!r||!canClient(req,r.client_id))return res.status(404).json({error:"결과물 없음"});
 db.prepare("DELETE FROM ai_results WHERE id=?").run(req.params.id);log(req,"delete","ai_result",req.params.id);res.json({ok:true});
});
app.get("/api/results/:id/export",auth,async(req,res)=>{
 try{
  const r=db.prepare("SELECT ar.*,c.name client_name FROM ai_results ar JOIN clients c ON c.id=ar.client_id WHERE ar.id=?").get(req.params.id);
  if(!r||!canClient(req,r.client_id))return res.status(404).json({error:"결과물 없음"});
  const result=JSON.parse(r.result_json||"{}"),fmt=String(req.query.format||"docx").toLowerCase(),safe=(r.title||"WORKI_AI_결과물").replace(/[\\/:*?"<>|]/g,"_").slice(0,80);
  if(fmt==="docx"){const b=await buildDocx(result,{role:r.role,clientName:r.client_name,title:r.title});res.setHeader("Content-Type","application/vnd.openxmlformats-officedocument.wordprocessingml.document");res.setHeader("Content-Disposition",`attachment; filename*=UTF-8''${encodeURIComponent(safe+".docx")}`);return res.send(b)}
  if(fmt==="pptx"){const b=await buildPptx(result,{role:r.role,clientName:r.client_name,title:r.title});res.setHeader("Content-Type","application/vnd.openxmlformats-officedocument.presentationml.presentation");res.setHeader("Content-Disposition",`attachment; filename*=UTF-8''${encodeURIComponent(safe+".pptx")}`);return res.send(b)}
  if(fmt==="csv"){const t=toCsv(result);res.setHeader("Content-Type","text/csv; charset=utf-8");res.setHeader("Content-Disposition",`attachment; filename*=UTF-8''${encodeURIComponent(safe+".csv")}`);return res.send(t)}
  if(fmt==="json"){res.setHeader("Content-Type","application/json; charset=utf-8");res.setHeader("Content-Disposition",`attachment; filename*=UTF-8''${encodeURIComponent(safe+".json")}`);return res.send(JSON.stringify(result,null,2))}
  const t=resultToPlain(result);res.setHeader("Content-Type","text/plain; charset=utf-8");res.setHeader("Content-Disposition",`attachment; filename*=UTF-8''${encodeURIComponent(safe+".txt")}`);res.send(t);
 }catch(e){res.status(500).json({error:e.message})}
});
app.get("/api/usage",auth,(req,res)=>{const c=tenantClause(req);res.json(db.prepare(`SELECT * FROM ai_usage WHERE ${c.sql} ORDER BY created_at DESC LIMIT 500`).all(...c.params))});
app.get("/api/audit",auth,admin,(req,res)=>res.json(db.prepare("SELECT * FROM audit_logs ORDER BY created_at DESC LIMIT 500").all()));

app.use((req,res)=>res.sendFile(path.join(__dirname,"public","index.html")));
app.listen(PORT,()=>console.log(`WORKI running on http://localhost:${PORT}`));
