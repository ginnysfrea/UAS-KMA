var KEY=[3, 1, 2, 1, 1, 1, 1, 0, 1, 0, 1, 1, 2, 0, 2, 1, 0, 2, 1, 2];
var ESS=[{"a": "Hadits shahih memenuhi lima syarat: sanad bersambung, perawi adil, perawi dhabith (teliti hafalan/catatan), tidak syadz, dan tidak mengandung ‘illah yang merusak.", "p": ["bersambung", "adil", "dhabith", "syadz", "illa"], "q": ["sanad", "perawi", "hafalan", "cacat"]}, {"a": "Hadits hasan: sanad bersambung, perawi adil, tidak syadz dan tidak ber-‘illah merusak, tetapi ketelitian (dhabith) perawinya berada di bawah perawi hadits shahih.", "p": ["dhabith", "bawah", "bersambung", "adil"], "q": ["syadz", "illa", "shahih", "maqbul"]}, {"a": "Hadits dhaif adalah hadits yang tidak memenuhi satu atau lebih syarat hadits shahih atau hasan, misalnya sanad terputus, perawi lemah hafalan, perawi majhul, atau ada cacat lain.", "p": ["tidak memenuhi", "shahih", "hasan"], "q": ["syarat", "terputus", "lemah", "majhul", "cacat"]}, {"a": "Tidak sama. Dhaif adalah riwayat yang tidak memenuhi syarat shahih/hasan, sedangkan maudhu’ adalah riwayat buatan yang dinisbatkan secara dusta kepada Nabi saw.", "p": ["tidak sama", "maudhu", "dusta", "dibuat"], "q": ["dhaif", "nabi", "syarat", "palsu"]}, {"a": "Hasan li dzatihi memenuhi syarat hasan berdasarkan kualitas sanad dan perawinya sendiri. Hasan li ghairihi awalnya dhaif ringan, lalu naik menjadi hasan karena ada jalur lain yang mendukung.", "p": ["dzatihi", "ghairihi", "jalur", "dhaif"], "q": ["sanad", "perawi", "hasan", "mendukung", "menguat"]}, {"a": "Takhrij hadits adalah kegiatan menelusuri hadits ke sumber aslinya, mengidentifikasi kitab periwayatnya, menghimpun jalur sanad dan redaksi matannya, serta membantu mengetahui kualitasnya berdasarkan penelitian ulama.", "p": ["menelusur", "sumber", "kitab"], "q": ["sanad", "matan", "jalur", "kualitas", "ulama"]}, {"a": "Kritik sanad adalah penelitian terhadap rangkaian perawi hadits. Yang diperiksa: kesinambungan sanad, keadilan perawi, ketelitian (dhabith), ada tidaknya syadz, dan ada tidaknya ‘illah yang merusak.", "p": ["rangkaian", "perawi", "bersambung", "adil", "dhabith"], "q": ["syadz", "illa", "penelitian", "kesinambungan"]}, {"a": "Kritik matan adalah penelitian terhadap isi/teks hadits: membandingkan redaksi antar jalur, memeriksa penambahan/pengurangan/kekeliruan, menilai pertentangan nyata dengan Al-Qur’an atau hadits yang lebih kuat, serta memeriksa konteks dan makna.", "p": ["isi", "teks", "banding", "redaksi"], "q": ["quran", "konteks", "makna", "penambahan", "pengurangan"]}, {"a": "Ittishal as-sanad adalah kesinambungan sanad: setiap perawi menerima hadits dari perawi sebelumnya. Diperiksa melalui tahun lahir/wafat, kesezamanan, hubungan guru-murid, bentuk periwayatan (haddatsana, akhbarana, ‘an), dan penilaian ulama.", "p": ["bersambung", "perawi", "guru", "murid"], "q": ["lahir", "wafat", "sezaman", "bertemu", "haddatsana", "ulama"]}, {"a": "Hadits syadz adalah riwayat perawi maqbul yang menyelisihi riwayat perawi lain yang lebih kuat. ‘Illah adalah cacat tersembunyi yang memengaruhi kualitas hadits meskipun sanad/matannya tampak baik.", "p": ["syadz", "menyelisihi", "illa", "tersembunyi"], "q": ["maqbul", "lebih kuat", "cacat", "kualitas"]}];
var HJ=['Waktu','Nama','NIM','Kode'],HN=['Nama','NIM','Kode','Skor PG (maks 50)'];
for(var i=1;i<=20;i++)HJ.push('PG '+i);
for(i=1;i<=10;i++){HJ.push('Essay '+i);HN.push('Essay '+i+' (maks 5)');}
HN.push('Total (maks 100)','Status');
function out(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);}
function tx(s){s=String(s);return /^[=+\-@]/.test(s)?"'"+s:s;}
function norm(s){return String(s).toLowerCase().replace(/[^a-z ]/g,'').replace(/\s+/g,' ');}
function mt(t,k){return k.length<=3?(' '+t+' ').indexOf(' '+k+' ')>=0:t.indexOf(k)>=0;}
function sc(e,v){var t=norm(v),a=e.p.filter(function(k){return mt(t,k)}).length,b=e.q.filter(function(k){return mt(t,k)}).length;return Math.round(5*(2*a+b)/(2*e.p.length+e.q.length)*10)/10;}
function sh(ss,name,h){var s=ss.getSheetByName(name)||ss.insertSheet(name);if(!s.getLastRow())s.appendRow(h);return s;}
function doPost(ev){
  var d=JSON.parse(ev.postData.contents),lock=LockService.getScriptLock();lock.waitLock(20000);
  try{
    var ss=SpreadsheetApp.getActiveSpreadsheet(),J=sh(ss,'Jawaban',HJ),N=sh(ss,'Penilaian',HN);
    var nim=String(d.nim).trim(),kode=String(d.kode).trim().toUpperCase(),i,idx=-1;
    var rows=N.getLastRow()>1?N.getRange(2,1,N.getLastRow()-1,16).getValues():[];
    for(i=0;i<rows.length;i++){if(String(rows[i][1])===nim){idx=i;break;}}
    if(d.a==='check'){
      if(idx<0||String(rows[idx][2]).toUpperCase()!==kode)return out({s:'nf'});
      var r=rows[idx];
      if(r[15]!=='Disetujui')return out({s:'wait'});
      var jr=J.getRange(idx+2,1,1,34).getValues()[0],es=0,p=[],e2=[];
      for(i=0;i<20;i++)p.push({a:jr[4+i],k:KEY[i]});
      for(i=0;i<10;i++){var s=Number(r[4+i])||0;es+=s;e2.push({t:jr[24+i],s:s,m:ESS[i].a});}
      return out({s:'ok',nama:r[0],pg:r[3],es:Math.round(es*10)/10,total:Math.round((Number(r[3])+es)*10)/10,p:p,e:e2});
    }
    if(idx>=0)return out(String(rows[idx][2]).toUpperCase()===kode?{ok:true}:{ok:false,msg:'NIM ini sudah pernah mengirim jawaban.'});
    var pg=0,L=[],P=d.p||[],A=d.e||[];
    for(i=0;i<20;i++){var c=P[i];L.push(c>=0?'ABCD'.charAt(c):'-');if(c===KEY[i])pg+=2.5;}
    J.appendRow([new Date(),tx(d.nama),"'"+nim,kode].concat(L,A.map(tx)));
    var sc2=ESS.map(function(x,k){return sc(x,A[k]||'');});
    var row=N.getLastRow()+1;
    N.appendRow([tx(d.nama),"'"+nim,kode,pg].concat(sc2,['=D'+row+'+SUM(E'+row+':N'+row+')','Menunggu']));
    N.getRange(row,16).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Menunggu','Disetujui'],true).build());
    return out({ok:true});
  }catch(err){return out({ok:false,msg:String(err)});}finally{lock.releaseLock();}
}
function setujuiSemua(){var N=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Penilaian'),n=N.getLastRow()-1;if(n>0)N.getRange(2,16,n,1).setValue('Disetujui');}
