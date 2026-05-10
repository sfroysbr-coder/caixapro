exports.handler = async (event) => {
  const h = {"Access-Control-Allow-Origin":"*","Access-Control-Allow-Headers":"Content-Type","Content-Type":"application/json"};
  if (event.httpMethod==="OPTIONS") return {statusCode:200,headers:h};
  if (event.httpMethod!=="POST") return {statusCode:405,headers:h,body:"Method Not Allowed"};
  try {
    const body = JSON.parse(event.body);
    const r = await fetch("https://api.anthropic.com/v1/messages",{
      method:"POST",
      headers:{"Content-Type":"application/json","x-api-key":process.env.ANTHROPIC_API_KEY,"anthropic-version":"2023-06-01"},
      body:JSON.stringify({model:"claude-sonnet-4-20250514",max_tokens:1000,messages:body.messages})
    });
    const data = await r.json();
    return {statusCode:200,headers:h,body:JSON.stringify(data)};
  } catch(e) {
    return {statusCode:500,headers:h,body:JSON.stringify({error:e.message})};
  }
};
