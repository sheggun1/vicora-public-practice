// Dedicated VICORA Public Practice API gateway.
function send(res,status,body){res.statusCode=status;res.setHeader("Content-Type","application/json; charset=utf-8");res.setHeader("Cache-Control","no-store");res.end(JSON.stringify(body))}
function message(v,fallback){if(typeof v==="string"&&v.trim())return v.trim();if(v&&typeof v==="object"){for(const k of ["message","error","details","description"])if(typeof v[k]==="string"&&v[k].trim())return v[k].trim();try{return JSON.stringify(v)}catch(_){}}return fallback}
module.exports=async function handler(req,res){
 const parts=(req.url||"/").split("?")[0].split("/").filter(Boolean).slice(2),path="/"+parts.join("/");
 const routes={"POST /register":"practice.register","POST /login":"practice.login","POST /logout":"practice.logout","GET /me":"practice.me","GET /catalog":"practice.catalog","GET /history":"practice.history","POST /start":"practice.start","POST /submit":"practice.submit","GET /admin/questions":"practice.admin.listQuestions","POST /admin/questions":"practice.admin.importQuestions"};
 const detail=path.match(/^\/history\/([^/]+)$/),action=detail&&req.method==="GET"?"practice.history.detail":routes[req.method+" "+path];
 if(!action)return send(res,404,{error:"Practice API endpoint not found."});
 const input=req.body&&typeof req.body==="object"?req.body:{},auth=String(req.headers.authorization||""),token=auth.startsWith("Bearer ")?auth.slice(7).trim():"";
 const payload={action,sessionToken:token};
 if(action==="practice.register"){payload.fullName=input.full_name;payload.email=input.email;payload.password=input.password}
 if(action==="practice.login")Object.assign(payload,input);
 if(action==="practice.start"){payload.examType=input.exam_type;payload.subject=input.subject;payload.questionCount=input.question_count}
 if(action==="practice.submit"){payload.attemptId=input.attempt_id;payload.answers=input.answers||{}}
 if(action==="practice.history.detail")payload.attemptId=decodeURIComponent(detail[1]);
 if(action.startsWith("practice.admin.")){const key=String(req.headers["x-vicora-practice-admin-key"]||"");if(key.length<32||key.length>200)return send(res,401,{error:"Enter a valid practice administrator key."});payload.adminKey=key;if(action==="practice.admin.importQuestions")payload.questions=input.questions}
 let endpoint;try{endpoint=new URL(process.env.VICORA_APPS_SCRIPT_URL||"")}catch(_){return send(res,503,{error:"VICORA_APPS_SCRIPT_URL is not configured in this Vercel project."})}
 if(endpoint.protocol!=="https:"||endpoint.hostname!=="script.google.com"||!/^\/macros\/s\/[^/]+\/exec$/.test(endpoint.pathname)||endpoint.search||endpoint.hash)return send(res,503,{error:"The configured Google Apps Script URL is invalid."});
 try{
  const upstream=await fetch(endpoint,{method:"POST",headers:{"Content-Type":"application/json","Accept":"application/json"},body:JSON.stringify(payload),redirect:"follow",signal:AbortSignal.timeout(25000)});
  const raw=await upstream.text();let envelope;try{envelope=JSON.parse(raw)}catch(_){return send(res,502,{error:"Google Apps Script returned an unreadable response."})}
  if(!upstream.ok||!envelope||envelope.ok!==true)return send(res,upstream.ok?400:502,{error:message(envelope&&envelope.error,"VICORA practice request failed.")});
  const data=envelope.data||{};let output=data;
  if(action==="practice.register"||action==="practice.login")output={token:data.sessionToken,user:{id:data.user.userId,full_name:data.user.fullName,email:data.user.email}};
  else if(action==="practice.me")output={user:{id:data.user.userId,full_name:data.user.fullName,email:data.user.email}};
  else if(action==="practice.start")output={attempt_id:data.attemptId,exam_type:data.examType,subject:data.subject,question_count:data.questionCount,questions:(data.questions||[]).map(q=>({id:q.id,year:q.year,topic:q.topic,question_text:q.questionText,options:q.options}))};
  else if(action==="practice.submit"||action==="practice.history.detail"){const result=data.result||data;output={result:{percent:result.percentage!==undefined?result.percentage:result.percent,score:result.score,total:result.total,correct_count:result.correctCount!==undefined?result.correctCount:result.correct_count,question_count:result.questionCount!==undefined?result.questionCount:result.question_count,exam_type:result.examType||result.exam_type,subject:result.subject},review:(data.review||[]).map(q=>({id:q.id,year:q.year,topic:q.topic,question_text:q.questionText||q.question_text,options:q.options,selected_option:q.selectedOption||q.selected_option,correct_option:q.correctOption||q.correct_option,correct:q.correct,explanation:q.explanation}))}}
  return send(res,200,output);
 }catch(_){return send(res,502,{error:"VICORA Apps Script is temporarily unavailable. Please try again."})}
};
