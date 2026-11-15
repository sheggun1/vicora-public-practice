// VICORA TEACHER API gateway. Keeps the Apps Script URL and provider credentials server-side.
function send(res,status,body){res.statusCode=status;res.setHeader("Content-Type","application/json; charset=utf-8");res.setHeader("Cache-Control","no-store");res.setHeader("X-Content-Type-Options","nosniff");res.end(JSON.stringify(body));}
function msg(v,fallback){if(typeof v==="string"&&v.trim())return v.trim();if(v&&typeof v==="object"){for(const k of ["message","error","details","description"])if(typeof v[k]==="string"&&v[k].trim())return v[k].trim();}return fallback;}
module.exports=async function handler(req,res){
 const parts=(req.url||"/").split("?")[0].split("/").filter(Boolean).slice(2),path="/"+parts.join("/");
 const routes={
  "POST /register":"teacher.register",
  "POST /login":"teacher.login",
  "POST /admin/login":"teacher.admin.login",
  "POST /logout":"teacher.logout",
  "GET /me":"teacher.me",
  "POST /documents/generate":"teacher.documents.generate",
  "POST /documents/save":"teacher.documents.save",
  "GET /documents":"teacher.documents.list",
  "GET /resources":"teacher.resources.list",
  "POST /resources/upload":"teacher.resources.upload",
  "POST /ai/chat":"teacher.ai.chat",
  "GET /admin/schools":"teacher.admin.schools",
  "GET /admin/ai-settings":"teacher.admin.aiSettings",
  "POST /admin/ai-settings":"teacher.admin.saveAiSettings"
 };
 const detail=path.match(/^\/(documents|resources)\/([A-Za-z0-9_-]+)$/);
 const action=detail&&req.method==="GET"?(detail[1]==="documents"?"teacher.documents.get":"teacher.resources.get"):routes[req.method+" "+path];
 if(!action)return send(res,404,{error:"VICORA Teacher API endpoint not found."});
 const input=req.body&&typeof req.body==="object"?req.body:{};
 const auth=String(req.headers.authorization||""),token=auth.startsWith("Bearer ")?auth.slice(7).trim():"";
 const payload={action,sessionToken:token};
 if(action==="teacher.register")Object.assign(payload,{fullName:input.full_name,email:input.email,password:input.password,schoolName:input.school_name,schoolState:input.school_state,schoolType:input.school_type,joinCode:input.join_code});
 if(action==="teacher.login")Object.assign(payload,{email:input.email,password:input.password});
 if(action==="teacher.admin.login")payload.adminKey=input.admin_key;
 if(action==="teacher.documents.generate")Object.assign(payload,{docType:input.doc_type,level:input.level,classLevel:input.class_level,subject:input.subject,term:input.term,week:input.week,topic:input.topic,subtopic:input.subtopic,duration:input.duration,questionCount:input.question_count,difficulty:input.difficulty,examFormat:input.exam_format,instructions:input.instructions,referenceIds:input.reference_ids||[]});
 if(action==="teacher.documents.get")payload.documentId=decodeURIComponent(detail[2]);
 if(action==="teacher.documents.save")Object.assign(payload,{documentId:input.document_id,contentHtml:input.content_html,contentText:input.content_text});
 if(action==="teacher.resources.get")payload.resourceId=decodeURIComponent(detail[2]);
 if(action==="teacher.resources.upload")Object.assign(payload,{fileName:input.file_name,mimeType:input.mime_type,fileBase64:input.file_base64,extractedText:input.extracted_text,description:input.description,targetSchoolId:input.target_school_id,visibility:input.visibility});
 if(action==="teacher.ai.chat")Object.assign(payload,{messages:input.messages||[],question:input.question,resourceIds:input.resource_ids||[]});
 if(action==="teacher.admin.saveAiSettings")Object.assign(payload,input);
 const query=new URLSearchParams((req.url||"").split("?")[1]||"");
 if(action==="teacher.documents.list"||action==="teacher.resources.list"){payload.limit=Math.min(100,Math.max(1,Number(query.get("limit"))||30));payload.type=query.get("type")||"";}
 let endpoint;try{endpoint=new URL(process.env.VICORA_TEACHER_APPS_SCRIPT_URL||"");}catch(_){return send(res,503,{error:"VICORA_TEACHER_APPS_SCRIPT_URL is not configured for this Teacher deployment."});}
 if(endpoint.protocol!=="https:"||endpoint.hostname!=="script.google.com"||!/^\/macros\/s\/[^/]+\/exec$/.test(endpoint.pathname)||endpoint.search||endpoint.hash)return send(res,503,{error:"The configured VICORA Teacher Apps Script URL is invalid."});
 try{
  const upstream=await fetch(endpoint,{method:"POST",headers:{"Content-Type":"application/json","Accept":"application/json"},body:JSON.stringify(payload),redirect:"follow",signal:AbortSignal.timeout(30000)});
  const raw=await upstream.text();let envelope;try{envelope=JSON.parse(raw);}catch(_){return send(res,502,{error:"VICORA Teacher backend returned an unreadable response."});}
  if(!upstream.ok||!envelope||envelope.ok!==true)return send(res,upstream.ok?400:502,{error:msg(envelope&&envelope.error,"VICORA Teacher request failed.")});
  return send(res,200,envelope.data||{});
 }catch(_){return send(res,502,{error:"VICORA Teacher backend is temporarily unavailable. Please try again."});}
};
