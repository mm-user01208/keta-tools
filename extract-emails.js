const T=[26280,26126,26122,26090,26037,25980,25976,25974,25947,25939,25892,25882,25881,25849,25823,25817,25815,25771,25754,25743,25705,25590,25588,25576,25569,25465,25457,25415,25359,25298,25253,25252,25242,25236,25167,25156,25085,25072,25064,25017,24993,24933,24888,24876,24843,24806,24793,24789,24787,24781,24780,24767,24726,24714,24689,24665,24629,24624,24609,24603];

(async()=>{
  var f=document.createElement("iframe");
  f.style.cssText="position:fixed;bottom:0;right:0;width:400px;height:300px;z-index:9999;border:2px solid #333";
  document.body.appendChild(f);
  var R=[];
  for(var i=0;i<T.length;i++){
    var id=T[i];
    console.log("処理中 "+(i+1)+"/"+T.length+" #"+id);
    await new Promise(function(r){f.onload=r;f.src="/issue/detail/"+id});
    await new Promise(function(r){setTimeout(r,2000)});
    var d=f.contentDocument;
    var btns=d.querySelectorAll("button");
    for(var j=0;j<btns.length;j++){
      if(btns[j].textContent.indexOf("\u30C1\u30B1\u30C3\u30C8\u306E\u8A73\u7D30\u3092\u898B\u308B")>=0){
        btns[j].click();break;
      }
    }
    await new Promise(function(r){setTimeout(r,1500)});
    var h=d.body.innerHTML;
    var m=h.match(/[\w.+-]+@[\w.-]+\.\w{2,}/g);
    var emails=[...new Set(m||[])].filter(function(x){return x!=="info@keta-travelsupport.net"});
    var email=emails[0]||"NOT_FOUND";
    var nm=d.body.innerText.match(/#\d+\s+([A-Z]+\s+[A-Z]+)\s+JP/);
    var name=nm?nm[1]:"";
    R.push(id+","+name+","+email);
    console.log(id+" "+name+" -> "+email);
  }
  document.body.removeChild(f);
  var csv="\u30C1\u30B1\u30C3\u30C8ID,\u540D\u524D,\u30E1\u30FC\u30EB\u30A2\u30C9\u30EC\u30B9\n"+R.join("\n");
  var blob=new Blob([csv],{type:"text/csv"});
  var a=document.createElement("a");
  a.href=URL.createObjectURL(blob);
  a.download="keta_cancel_emails.csv";
  a.click();
  console.log("--- CSV ---");
  console.log(csv);
})();
