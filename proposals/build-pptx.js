const pptxgen = require('pptxgenjs');
const p = new pptxgen();
p.layout = 'LAYOUT_WIDE';           // 13.333 x 7.5
p.author = '米倉まな';
p.title  = '米倉まな｜アポプラスステーション様 ご提案 導入';

const PAPER='F3EADB', INK='2B2620', INK2='5A5249', INK3='8B8175', RULE='D2C4AE',
      SHU='9C5446', SHU_L='C97B68', MIDORI='5E7060', CARD='FBF6EC',
      DFG='F3ECE0', DFG2='BDB2A2', DRULE='4A423A';
const SERIF='游明朝', SANS='游ゴシック';
const M=0.75, W=13.333, CW=W-M*2;

function bg(s,dark){ s.background={color: dark?INK:PAPER}; }
function eyebrow(s,num,label,dark){
  s.addShape(p.ShapeType.rect,{x:M,y:0.615,w:0.26,h:0.014,fill:{color:dark?SHU_L:SHU},line:{type:'none'}});
  const runs=[];
  if(num) runs.push({text:num+'   ',options:{fontFace:SERIF,color:dark?SHU_L:SHU,bold:true,charSpacing:2}});
  runs.push({text:label,options:{color:dark?DFG2:INK3,charSpacing:3}});
  s.addText(runs,{x:1.09,y:0.44,w:8,h:0.34,fontSize:10,fontFace:SANS,isTextBox:true,margin:0,valign:'middle'});
}
function title(s,t,dark,size){
  s.addText(t,{x:M,y:0.92,w:CW,h:0.85,fontSize:size||28,fontFace:SERIF,bold:true,
    color:dark?DFG:INK,isTextBox:true,margin:0,valign:'middle',charSpacing:1});
}
function foot(s,time,dark){
  s.addText('米倉まな｜アポプラスステーション様 ご提案',{x:M,y:6.93,w:8,h:0.3,fontSize:9,fontFace:SANS,
    color:dark?DFG2:INK3,isTextBox:true,margin:0,valign:'middle',charSpacing:1});
  s.addText(time,{x:W-M-3.5,y:6.93,w:3.5,h:0.3,fontSize:9,fontFace:SANS,color:dark?DFG2:INK3,
    isTextBox:true,margin:0,valign:'middle',align:'right'});
}
function src(s,t,dark,y){
  s.addText('出典：'+t,{x:M,y:y||6.52,w:CW,h:0.36,fontSize:7.5,fontFace:SANS,
    color:dark?DFG2:INK3,isTextBox:true,margin:0,valign:'top',lineSpacing:11});
}
function hr(s,x,y,w,dark){
  s.addShape(p.ShapeType.rect,{x,y,w,h:0.008,fill:{color:dark?DRULE:RULE},line:{type:'none'}});
}
function card(s,x,y,w,h,tag,k,v){
  s.addShape(p.ShapeType.roundRect,{x,y,w,h,rectRadius:0.05,fill:{color:CARD},line:{color:RULE,width:0.75}});
  s.addText(tag,{x:x+0.28,y:y+0.22,w:w-0.56,h:0.26,fontSize:9,fontFace:SANS,bold:true,color:SHU,
    charSpacing:2,isTextBox:true,margin:0,valign:'middle'});
  s.addText(k,{x:x+0.28,y:y+0.52,w:w-0.56,h:0.5,fontSize:15,fontFace:SERIF,bold:true,color:INK,
    isTextBox:true,margin:0,valign:'top',lineSpacing:21});
  s.addText(v,{x:x+0.28,y:y+1.08,w:w-0.56,h:h-1.3,fontSize:10.5,fontFace:SANS,color:INK2,
    isTextBox:true,margin:0,valign:'top',lineSpacing:17});
}
function quote(s,x,y,w,h,q1,runs,tint){
  const c = tint||MIDORI;
  s.addShape(p.ShapeType.roundRect,{x,y,w,h,rectRadius:0.05,fill:{color:CARD},line:{color:c,width:0.75}});
  s.addText(q1,{x:x+0.3,y:y+0.22,w:w-0.6,h:0.26,fontSize:9,fontFace:SANS,bold:true,color:c,
    charSpacing:2,isTextBox:true,margin:0,valign:'middle'});
  s.addText(runs,{x:x+0.3,y:y+0.56,w:w-0.6,h:h-0.8,fontSize:11,fontFace:SANS,color:INK2,
    isTextBox:true,margin:0,valign:'top',lineSpacing:19});
}
function bullets(s,x,y,w,items,dark,size){
  const runs=[];
  items.forEach((it,i)=>{
    (Array.isArray(it)?it:[{text:it}]).forEach((r,j)=>{
      runs.push({text:r.text,options:Object.assign({
        bullet:(j===0)?{code:'2022'}:false, color:r.b?(dark?DFG:INK):(dark?DFG2:INK2), bold:!!r.b
      }, j===0?{}:{})});
    });
    runs[runs.length-1].options.breakLine = (i!==items.length-1);
    runs[runs.length-1].options.paraSpaceAfter = 10;
  });
  s.addText(runs,{x,y,w,h:4.2,fontSize:size||11,fontFace:SANS,isTextBox:true,margin:0,
    valign:'top',lineSpacing:19});
}

/* ---------- 1 表紙 ---------- */
let s=p.addSlide(); bg(s);
s.addShape(p.ShapeType.rect,{x:M,y:0.615,w:0.26,h:0.014,fill:{color:SHU},line:{type:'none'}});
s.addText('アポプラスステーション株式会社 御中',{x:1.09,y:0.44,w:8,h:0.34,fontSize:11,fontFace:SANS,
  color:SHU,charSpacing:2,isTextBox:true,margin:0,valign:'middle'});
s.addText('よねくら　まな',{x:M,y:1.95,w:6,h:0.3,fontSize:11,fontFace:SANS,color:INK3,charSpacing:5,
  isTextBox:true,margin:0,valign:'middle'});
s.addText('米倉まな',{x:M,y:2.35,w:8,h:0.95,fontSize:38,fontFace:SERIF,bold:true,color:INK,
  charSpacing:3,isTextBox:true,margin:0,valign:'middle'});
s.addText('ご提案の前に、\n私が何者で、何を変えたいのかを。',{x:M,y:3.62,w:9,h:1.2,fontSize:20,
  fontFace:SERIF,color:INK2,isTextBox:true,margin:0,valign:'top',lineSpacing:34});
const pills=['認定鍼灸師（開業11年目）','はりきゅう処ここちめいど 院長','オンラインサロン ここちめいど 主宰','ここちはり 代表','武蔵野大学大学院 人間学 修士'];
let px=M;
pills.forEach(t=>{
  const w=t.length*0.125+0.42;
  s.addShape(p.ShapeType.roundRect,{x:px,y:5.35,w,h:0.42,rectRadius:0.21,fill:{color:CARD},line:{color:RULE,width:0.75}});
  s.addText(t,{x:px,y:5.35,w,h:0.42,fontSize:9.5,fontFace:SANS,color:INK2,align:'center',
    isTextBox:true,margin:0,valign:'middle'});
  px+=w+0.16;
});
foot(s,'所要 約5分');
s.addNotes('本日はお時間をいただきありがとうございます。提案に入る前に、私が何者で、何を変えたいのかを5分だけお話しさせてください。');

/* ---------- 2 原点（濃色） ---------- */
s=p.addSlide(); bg(s,true); eyebrow(s,'01','ORIGIN',true);
title(s,'私は2013年ごろまで、患者の側にいました',true);
[['01',[{t:'精神疾患の当事者として、'},{t:'約8年間さまよいました。',b:1}]],
 ['02',[{t:'病院にも通いました。よくならない時期が、とても長く続きました。'}]],
 ['03',[{t:'最後に私を助けたのが、鍼灸でした。いまこの仕事をしているのは、そのためです。'}]]
].forEach((f,i)=>{
  const y=2.2+i*0.85;
  s.addText(f[0],{x:M,y,w:0.5,h:0.4,fontSize:14,fontFace:SERIF,bold:true,color:SHU_L,isTextBox:true,margin:0,valign:'middle'});
  s.addText(f[1].map(r=>({text:r.t,options:{color:r.b?DFG:DFG2,bold:!!r.b}})),
    {x:M+0.6,y,w:5.4,h:0.75,fontSize:11.5,fontFace:SANS,isTextBox:true,margin:0,valign:'middle',lineSpacing:19});
});
s.addText('「もっと早く、\n鍼灸に出会いたかった」',{x:7.2,y:2.15,w:5.38,h:1.5,fontSize:24,fontFace:SERIF,bold:true,
  color:DFG,isTextBox:true,margin:0,valign:'top',lineSpacing:40});
s.addText('患者さんから、いちばん多く言われる言葉です。この「もっと早く」をなくすことが、私の仕事だと思っています。',
  {x:7.2,y:3.85,w:5.38,h:1.0,fontSize:11,fontFace:SANS,color:DFG2,isTextBox:true,margin:0,valign:'top',lineSpacing:19});
foot(s,'0:20 → 0:55',true);
s.addNotes('まず、私自身の話をさせてください。私は2013年ごろまで、精神疾患の当事者でした。約8年間さまよって、病院にも通って、よくならない時期がとても長く続きました。最後に私を助けたのが鍼灸です。いまこの仕事をしているのは、そのためです。そして患者さんからいちばん多く言われるのが、「もっと早く鍼灸に出会いたかった」という言葉です。この「もっと早く」をなくすことが、私の仕事だと思っています。');

/* ---------- 3 プロフィール ---------- */
s=p.addSlide(); bg(s); eyebrow(s,'02','PROFILE');
title(s,'臨床・育成・研究の3本を、11年まわしてきました');
card(s,M,1.9,3.78,2.5,'臨床','はりきゅう処ここちめいど','三重県四日市市。2015年開院、夜22時まで。心の不調を主訴とする患者を中心に診ています。2026年9月に町田院を開院予定。');
card(s,4.78,1.9,3.78,2.5,'育成・コミュニティ','ここちめいど／ここちはり','傾聴を学ぶ鍼灸師のオンラインサロンを2020年から主宰。2025年からは施術者育成事業を開始し、感情・心・身体・技術の4軸で指導。');
card(s,8.81,1.9,3.78,2.5,'研究','学会発表・産学連携','全日本鍼灸学会・日本うつ病学会で継続的に発表し、日本精神神経学会にも登壇。理化学研究所・東京有明医療大学・玉川大学・ケアクルとの共同研究に参画。');
hr(s,M,4.85,CW);
const stats=[['11年目','鍼灸院 開業'],['20,000人+','のべ臨床患者数'],['3学会','発表・登壇'],['修士','武蔵野大学大学院 人間学']];
stats.forEach((st,i)=>{
  const x=M+i*(CW/4);
  s.addText(st[0],{x,y:5.05,w:CW/4-0.2,h:0.6,fontSize:26,fontFace:SERIF,bold:true,color:INK,
    isTextBox:true,margin:0,valign:'middle'});
  s.addText(st[1],{x,y:5.68,w:CW/4-0.2,h:0.3,fontSize:9.5,fontFace:SANS,color:INK3,charSpacing:1,
    isTextBox:true,margin:0,valign:'middle'});
});
foot(s,'0:55 → 1:30');
s.addNotes('三重で鍼灸院を11年。のべ2万人以上を診てきました。臨床だけでなく、鍼灸師の育成と、学会での研究発表も続けています。全日本鍼灸学会と日本うつ病学会で発表し、日本精神神経学会にも登壇しました。心の不調を扱う鍼灸師として、現場・人・データの3つを持っているのが私の立ち位置です。');

/* ---------- 3 強み ---------- */
s=p.addSlide(); bg(s); eyebrow(s,'03','STRENGTH');
title(s,'企業の新規事業に、専門家として入ってきました');
card(s,M,1.9,3.78,2.65,'パパゲーノ','アンバサダー認定 第一号','就労継続支援B型「パパゲーノ Work & Recovery」への発注を通じ、精神障害のある方が担う事業に伴走。認定制度の第一号企業です。');
card(s,4.78,1.9,3.78,2.65,'ファンコミュニケーションズ','ThirdTalk（サードトーク）','「病院やカウンセリングに行く、その一歩手前」を想定したオンライン対話サービス。企画段階からご相談をいただき、立ち上げを見守ってきました。');
card(s,8.81,1.9,3.78,2.65,'stand up','アミューズメント事業','新規事業の立ち上げについてご相談をいただき、現在もお取引が続いています。医療以外の業種でも、人の状態を扱う事業の設計に入ってきました。');
quote(s,M,4.78,CW,1.5,'共通してやってきたこと',[
  {text:'心の不調を抱えた人が、実際にどこで止まり、どこで動くのか。その一次情報を事業の言葉に翻訳するのが私の役割です。なかでもThirdTalkは、このあとお話しする'},
  {text:'「病院の一歩手前」という発想を、そのまま事業にしたもの',options:{bold:true,color:INK}},
  {text:'でした。'}]);
src(s,'株式会社ファンコミュニケーションズ ニュースリリース（2022.8.26）／株式会社パパゲーノ プレスリリース',false,6.42);
foot(s,'1:30 → 2:10');
s.addNotes('もうひとつの顔が、企業の新規事業への伴走です。パパゲーノでは認定第一号として、就労継続支援の事業に発注する側で関わりました。ファンコミュニケーションズさんのThirdTalkは、病院やカウンセリングに行く一歩手前の人のためのオンライン対話サービスで、企画段階からご相談をいただき、立ち上げを見守ってきました。stand upさんはアミューズメントの会社で、新規事業のご相談から、いまもお取引が続いています。');

/* ---------- 4 課題感（濃色） ---------- */
s=p.addSlide(); bg(s,true); eyebrow(s,'04','ISSUE',true);
title(s,'これは私の実感であり、社会の数字でもあります',true);
const LX=M, LW=5.5, RX=6.9, RW=5.68;
s.addText('社会で起きていること',{x:LX,y:1.78,w:LW,h:0.28,fontSize:9.5,fontFace:SANS,bold:true,color:SHU_L,charSpacing:3,isTextBox:true,margin:0,valign:'middle'});
s.addText('私が現場で見ていること',{x:RX,y:1.78,w:RW,h:0.28,fontSize:9.5,fontFace:SANS,bold:true,color:'93AC93',charSpacing:3,isTextBox:true,margin:0,valign:'middle'});
const rows=[
  [[{text:'約600万人　',f:1},{text:'精神疾患を持つ患者数。国は2024年、厚生労働白書で初めて「こころの健康」をテーマに据えました。'}],
   'のべ20,000人を診てきて、心の不調を主訴に来られる方が明らかに増えました。'],
  [[{text:'2028年4月　',f:1},{text:'50人未満の事業場にもストレスチェックが義務化。現在の実施率は6割弱です。'}],
   '気づいても、次に行く場所がない。精神科は予約が埋まり、新規の受付を止める医療機関が出ています。'],
  [[{text:'年7.6兆円　',f:1},{text:'メンタル不調による経済損失。うち7.3兆円は「出勤しているのに力が出ない」分です。'}],
   '休職や通院の手前で止まっている方が、いちばん多い。この層は、いま誰の担当でもありません。']
];
rows.forEach((r,i)=>{
  const y=2.14+i*1.06;
  hr(s,LX,y-0.12,LW,true); hr(s,RX,y-0.12,RW,true);
  s.addText(r[0].map(t=>({text:t.text,options:t.f?{fontFace:SERIF,fontSize:15,bold:true,color:SHU_L}:{color:DFG}})),
    {x:LX,y,w:LW,h:0.9,fontSize:10.5,fontFace:SANS,isTextBox:true,margin:0,valign:'top',lineSpacing:18});
  s.addText(r[1],{x:RX,y,w:RW,h:0.9,fontSize:10.5,fontFace:SANS,color:DFG2,isTextBox:true,margin:0,valign:'top',lineSpacing:18});
});
s.addText('だから、未病の段階で止めておきたい。',{x:LX,y:5.6,w:CW,h:0.6,fontSize:23,fontFace:SERIF,bold:true,
  color:SHU_L,isTextBox:true,margin:0,valign:'middle'});
src(s,'厚生労働省 患者調査／令和6年版 厚生労働白書／改正労働安全衛生法（2025年5月公布・2028年4月施行）／メンタルヘルス不調による経済損失の試算（労働者27,507人調査）',true,6.45);
foot(s,'2:10 → 3:00',true);
s.addNotes('ここからは、私が現場で感じていることをお話しします。左が社会の数字、右が私が院で見ていることです。精神疾患の患者は約600万人、国も2024年に初めて白書のテーマにこころの健康を据えました。私の院でも、心の不調を主訴に来られる方が明らかに増えています。2028年からは50人未満の会社にもストレスチェックが義務化されます。ただ、検査で気づいても次に行く場所がない。病院はもう新規を受けきれていません。そして損失の7.6兆円のうち7.3兆円は、出勤しているのに力が出ない人の分です。この層が、いま誰の担当にもなっていない。だから私は、未病の段階で止めたいと考えています。');

/* ---------- 5 打ち手 ---------- */
s=p.addSlide(); bg(s); eyebrow(s,'05','ACTION');
title(s,'新規事業「ねここち」を立ち上げます');
bullets(s,M,2.05,7.3,[
  [{text:'ねここち',b:1},{text:'／頭オフ事業。考えごとで止まらない頭を、強制的にオフにする時間を売ります。いま立ち上げを準備しています。'}],
  [{text:'施術は'},{text:'鍼灸師の国家資格を持つ人',b:1},{text:'が行います。医療の知識を持つ人が、治療の手前にいる層に関わる形です。'}],
  [{text:'「治療」と身構えずに入れる入口をつくり、'},{text:'必要な人だけを鍼灸や医療につなぐ',b:1},{text:'。未病の層に、継続的に関わり続けるための事業です。'}]
]);
s.addShape(p.ShapeType.roundRect,{x:8.45,y:1.95,w:4.13,h:2.5,rectRadius:0.05,fill:{color:'F3E3DE'},line:{color:SHU,width:0.75}});
s.addText('1店舗あたり',{x:8.45,y:2.2,w:4.13,h:0.3,fontSize:9.5,fontFace:SANS,bold:true,color:SHU,
  charSpacing:3,align:'center',isTextBox:true,margin:0,valign:'middle'});
s.addText([{text:'月次利益 30',options:{}},{text:'万円',options:{fontSize:17}}],
  {x:8.45,y:2.55,w:4.13,h:0.8,fontSize:34,fontFace:SERIF,bold:true,color:INK,align:'center',
   isTextBox:true,margin:0,valign:'middle'});
s.addText('想定値。多店舗展開を前提に設計しています。',
  {x:8.75,y:3.45,w:3.53,h:0.8,fontSize:9.5,fontFace:SANS,color:INK2,align:'center',
   isTextBox:true,margin:0,valign:'top',lineSpacing:16});
foot(s,'3:00 → 3:35');
s.addNotes('いま「ねここち」という、頭をオフにする事業の立ち上げを準備しています。頭が考えごとで止まらない人に、強制的にオフにする時間を売る事業です。施術は鍼灸師の国家資格を持つ人が行います。医療の知識を持つ人が、治療の手前にいる層に関わる形です。治療と身構えずに入れる入口をつくって、必要な人だけを鍼灸や医療につなぐ。1店舗あたり月次利益30万円を想定しています。');

/* ---------- 6 先行事例 ---------- */
s=p.addSlide(); bg(s); eyebrow(s,'06','CASE');
title(s,'「企業 × 鍼灸」は、すでに成立しています');
bullets(s,M,1.95,6.6,[
  [{text:'睡眠製品のブレインスリープ社は、'},{text:'睡眠専門の鍼灸マッサージ「コンディショニングスタジオ」',b:1},{text:'を自社ブランドで運営。枕の売り場ではなく、独立した鍼灸院です。'}],
  [{text:'二子玉川・南青山・'},{text:'有楽町マルイ',b:1},{text:'・東京ミッドタウン八重洲の'},{text:'4店舗',b:1},{text:'。商業施設側である丸井グループも、自らプレスリリースを出しています。'}],
  [{text:'オープン以来'},{text:'累計1万人以上',b:1},{text:'が利用。自社研究では、1度の施術で眠りの深さが'},{text:'35%',b:1},{text:'、疲労回復度が'},{text:'59%',b:1},{text:'向上したと公表。'}],
  [{text:'スタッフは全員が国家資格保有者で、自社資格「スリーププランナー」を取得。'},{text:'定額プラン',b:1},{text:'で、継続して通う前提が設計されています。'}]
],false,10.5);
quote(s,7.7,1.95,4.88,3.5,'ここから私が受け取ったこと',[
  {text:'製品を売って終わりにせず、施術という接点で人と関わり続ける。効果を数字で示し、資格者を自社で育て、定額で通ってもらう。——私がやろうとしている構造と、ほぼ同じです。\n\n違いは入口です。'},
  {text:'あの会社は「睡眠」から入った。私は「こころ」から入ります。',options:{bold:true,color:INK}}]);
src(s,'株式会社丸井グループ プレスリリース（2024.11.15）／株式会社ブレインスリープ プレスリリース／ブレインスリープ コンディショニングスタジオ 公式サイト',false,6.4);
foot(s,'3:35 → 4:15');
s.addNotes('企業と鍼灸の組み合わせは、もう絵空事ではありません。枕で知られるブレインスリープさんは、睡眠専門の鍼灸院を自社ブランドで4店舗運営しています。有楽町マルイのように商業施設の中にも入っていて、丸井グループ側もプレスリリースを出しています。累計1万人以上が利用し、施術の効果も数字で公表し、定額プランで通い続けてもらう設計です。製品を売って終わりにせず、施術で人と関わり続けている。私がやろうとしている構造とほぼ同じで、違いは入口だけです。あの会社は睡眠から入った。私はこころから入ります。');

/* ---------- 7 当社の課題（濃色） ---------- */
s=p.addSlide(); bg(s,true); eyebrow(s,'07','OUR LIMIT',true);
title(s,'正直に言うと、私1人で抱えられる量に天井があります',true);
[['01','臨床も、育成も、研究も、いまは私に紐づいている（属人）'],
 ['02','店舗は増やせても、施術者と運営を担える人は増やせていない'],
 ['03','課題の大きさに対して、届けられる人数が圧倒的に足りない']].forEach((f,i)=>{
  const y=2.25+i*0.85;
  s.addText(f[0],{x:M,y,w:0.5,h:0.4,fontSize:14,fontFace:SERIF,bold:true,color:SHU_L,isTextBox:true,margin:0,valign:'middle'});
  s.addText(f[1],{x:M+0.6,y,w:5.4,h:0.7,fontSize:11.5,fontFace:SANS,color:DFG,isTextBox:true,margin:0,valign:'middle',lineSpacing:19});
 });
s.addText('仕組みと、人と、\n座組みが要る。',{x:7.2,y:2.2,w:5.38,h:1.5,fontSize:26,fontFace:SERIF,bold:true,
  color:DFG,isTextBox:true,margin:0,valign:'top',lineSpacing:42});
s.addText('ここは私が最も苦手としてきた領域であり、御社が最も強い領域だと考えています。',
  {x:7.2,y:3.95,w:5.38,h:0.9,fontSize:11,fontFace:SANS,color:DFG2,isTextBox:true,margin:0,valign:'top',lineSpacing:19});
foot(s,'4:15 → 4:35',true);
s.addNotes('正直に申し上げると、いまはすべてが私に紐づいていて、1人で抱えられる量に天井があります。仕組みと人と座組みが足りない。ここは御社が最も強い領域だと考えています。');

/* ---------- 8 ご相談 ---------- */
s=p.addSlide(); bg(s); eyebrow(s,'08','ASK');
title(s,'だから、アポプラスステーション様と話したい');
[['01','未病領域の新規事業を、共同で検討させてください','東洋医学 × 御社の事業アセット。入口となるサービス設計から、実証までを一緒に。'],
 ['02','人と組織の知見をお借りしたい','施術者の採用・育成・配置。属人から仕組みへ移すために、御社の人材領域の知見が必要です。'],
 ['03','実証の場をご一緒したい','法人向けの健康支援など、未病層に届く場での実証。効果は学会発表の形で残し、根拠として積み上げます。']
].forEach((a,i)=>{
  const y=1.95+i*1.0;
  if(i>0) hr(s,M,y-0.16,CW);
  s.addText(a[0],{x:M,y,w:0.6,h:0.4,fontSize:16,fontFace:SERIF,bold:true,color:SHU,isTextBox:true,margin:0,valign:'middle'});
  s.addText(a[1],{x:M+0.75,y:y-0.02,w:11,h:0.36,fontSize:14,fontFace:SERIF,bold:true,color:INK,isTextBox:true,margin:0,valign:'middle'});
  s.addText(a[2],{x:M+0.75,y:y+0.36,w:11,h:0.36,fontSize:10.5,fontFace:SANS,color:INK2,isTextBox:true,margin:0,valign:'middle'});
});
quote(s,M,5.1,CW,1.3,'私の目的',[{text:'生きづらさを抱えた人を、治療が必要になる手前で減らしたい。そのために、鍼灸の外側にある方法も含めて手を尽くします。'}],SHU);
foot(s,'4:35 → 5:00');
s.addNotes('お願いしたいことは3つです。未病領域の新規事業の共同検討、人材・組織の知見、そして実証の場。治療が必要になる手前で、生きづらさを抱えた人を減らしたい。ぜひご一緒させてください。');

/* ---------- 9 参考資料 ---------- */
s=p.addSlide(); bg(s); eyebrow(s,'','APPENDIX ／ 参考');
title(s,'公開されている実績・共同研究');
[['共同研究','開業鍼灸院でのメンタルヘルスの可能性を検証する共同研究','ここちめいど／東京有明医療大学／理化学研究所／ケアクル　― PR TIMES'],
 ['共同研究','鍼灸症例データベース構築プロジェクト','ケアクル／東京有明医療大学 松浦悠人／玉川大学 柴田健一／ここちめいど 米倉まな　― PR TIMES（2025.4）'],
 ['企業連携','オンライン対話サービス「ThirdTalk」特別対談番組に出演','株式会社ファンコミュニケーションズ ニュースリリース（2022.8.26）'],
 ['社会貢献','パパゲーノ・アンバサダー認定 第一号（合同会社ここちめいど）','株式会社パパゲーノ プレスリリース'],
 ['寄稿','『精神看護』特集「鍼灸は、なぜ「うつ」に効くのだろうか」','医学書院　Vol.28 No.3（2025.5）／ レポート「障害と美の多様性」Vol.28 No.5（2025.9）'],
 ['学会','全日本鍼灸学会（2024・2025・2026）／ 日本うつ病学会（ポスター発表）／ 日本精神神経学会 鍼灸シンポジウム シンポジスト','うつと不安症状を有する患者の実態調査 第3報・第4報 ほか　― 日本精神神経学会は2026.6 登壇']
].forEach((r,i)=>{
  const y=1.95+i*0.8;
  hr(s,M,y-0.12,CW);
  s.addText(r[0],{x:M,y:y+0.02,w:1.1,h:0.28,fontSize:9,fontFace:SANS,bold:true,color:SHU,charSpacing:2,isTextBox:true,margin:0,valign:'middle'});
  s.addText(r[1],{x:M+1.25,y:y-0.02,w:10.58,h:0.3,fontSize:11,fontFace:SANS,color:INK,isTextBox:true,margin:0,valign:'middle'});
  s.addText(r[2],{x:M+1.25,y:y+0.3,w:10.58,h:0.3,fontSize:9,fontFace:SANS,color:INK3,isTextBox:true,margin:0,valign:'middle'});
});
hr(s,M,1.95+6*0.8-0.12,CW);
s.addText('参考資料（本編5分には含みません）',{x:W-M-4,y:6.93,w:4,h:0.3,fontSize:9,fontFace:SANS,color:INK3,
  isTextBox:true,margin:0,valign:'middle',align:'right'});
s.addText('米倉まな｜アポプラスステーション様 ご提案',{x:M,y:6.93,w:8,h:0.3,fontSize:9,fontFace:SANS,color:INK3,
  isTextBox:true,margin:0,valign:'middle',charSpacing:1});
s.addNotes('本編には含めません。実績の裏取りを求められたときに、共同研究や企業連携の公開情報としてお見せする用です。');

p.writeFile({fileName:'米倉まな_アポプラスステーション様ご提案_導入.pptx'}).then(f=>console.log('written',f));
