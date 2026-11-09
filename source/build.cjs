const PptxGenJS = require("pptxgenjs");
const data = require("./scenes.json");
const deck = new PptxGenJS();
deck.layout = "LAYOUT_WIDE";
deck.rtlMode = true;
deck.title = "بوابة وادي التنانين";
deck.subject = "تجربة تعليمية في ورشة Claude";
deck.lang = "ar-SA";
const S = deck.ShapeType;
const W = 13.333, H = 7.5;
const C = { forest:"193B34", deep:"102A25", ivory:"FAF8F2", sand:"DBC09A", mint:"83E0C1", pale:"E5F4EC", line:"517369", ink:"172621", muted:"526660", white:"FFFFFF" };
const fontFace = "Arial";
function rect(slide,x,y,w,h,color,r=0) {
  slide.addShape(r ? S.roundRect : S.rect,{x,y,w,h,rectRadius:r,radius:r,line:{transparency:100},fill:{color}});
}
function tx(slide,t,x,y,w,h,opts={}) {
  slide.addText(t,{x,y,w,h,fontFace,fontSize:20,rtlMode:true,align:"right",valign:"mid",margin:0,breakLine:false,color:C.ink,lang:"ar-SA",...opts});
}
function path(slide,points,color,width=1.1) {
  for(let i=1;i<points.length;i++){
    const [a,b]=points[i-1], [c,d]=points[i];
    slide.addShape(S.line,{x:Math.min(a,c),y:Math.min(b,d),w:Math.abs(a-c),h:Math.abs(b-d),flipV:(c-a)*(d-b)<0,line:{color,width,transparency:12}});
  }
}
function card(slide,title,body,x,y,w,dark){
  rect(slide,x,y,w,1.55,dark?"254C43":"FFFFFF");
  rect(slide,x+w-.08,y,.08,1.55,dark?C.sand:C.forest);
  tx(slide,title,x+.18,y+.14,w-.44,.55,{fontSize:21,bold:true,color:dark?C.ivory:C.forest});
  tx(slide,body,x+.18,y+.82,w-.44,.45,{fontSize:14,color:dark?"D4E5DC":C.muted});
}
const descriptions=[
  [],
  ["أعطال التبريد والخدمة","رفع الأسعار داخل المحاكاة","دخول منتجع العنقاء"],
  [],
  ["حدّد المطلوب من خطاب التحدي","اطلب طريقة الحساب والقيود","حوّل الفكرة إلى وظيفة قابلة للاختبار"],
  ["استعراض الإقامة وحساب التكلفة","رقم مرجعي وحالة للطلب","أسئلة ثم توصية"],
  ["المراجعة كشفت تكرار تجربة قائمة","وظيفة مستقلة تسهّل الوصول"],
  ["ما الذي تغير؟ وما مصدر الدليل؟","ماذا يستطيع النموذج إظهاره؟","جرّب المسار وسجّل ما تعثر"],
  []
];
const labels=["الافتتاح","التكليف","الأدلة","التعلّم مع Claude","الحل","المراجعة","مهارات قابلة للنقل","ختام الورشة"];
for(const [i,[title,subtitle,items,seconds,notes]] of data.entries()){
  const slide=deck.addSlide();
  const dark=[0,3,4,5].includes(i), fg=dark?C.ivory:C.ink, soft=dark?"D5E3DB":C.muted;
  slide.background={color:dark?C.deep:i===7?"F3FBF6":C.ivory};
  rect(slide,12.66,.64,.06,1.45,dark?C.sand:C.forest);
  tx(slide,labels[i],8,.57,4.46,.31,{color:dark?C.sand:C.forest,fontSize:13,bold:true});
  tx(slide,title,.68,1.02,11.75,.84,{color:fg,fontSize:i===0?49:i===7?40:34,bold:true});
  tx(slide,subtitle,.7,1.95,11.72,.6,{color:soft,fontSize:20});
  if(i===0){
    if(process.env.NOON_TEAM) tx(slide,process.env.NOON_TEAM,7.05,5.75,5.35,.55,{fontSize:14,color:soft});
    for(let j=0;j<4;j++){
      const o=j*.31;
      path(slide,[[.65,6.45-o],[2.7,4.2-o],[3.38,4.71-o],[4.7,3.18-o],[6.55,5.48-o],[7.1,4.91-o]],j===0?C.sand:C.line,1.6);
    }
    tx(slide,"تحليل المشكلة · بناء الحل · مراجعة النتيجة",7.25,4.3,5.15,.72,{fontSize:22,color:C.sand,bold:true});
    tx(slide,"ورشة Claude لدى نون",8.3,5.15,4.1,.45,{fontSize:17,color:soft});
  }else if(i===2){
    if(process.env.NOON_SIMPLE_CHART){
      for(const [k,[year,value]] of [["2025",72.03],["2026",60.52]].entries()){
        const x=1.2+k*3.7,h=value/100*2.55;
        rect(slide,x,5.85-h,2.2,h,k===0?C.forest:"52766B");
        tx(slide,value.toFixed(2)+"%",x,5.42-h,2.2,.34,{align:"center",rtlMode:false,bold:true});
        tx(slide,year,x,5.96,2.2,.32,{align:"center",rtlMode:false,fontSize:15});
      }
      path(slide,[[.85,5.85],[8.5,5.85]],C.line,1);
    }else slide.addChart(deck.ChartType.bar,[{name:"الإشغال",labels:["2025","2026"],values:[.7203,.6052]}],{
      x:1,y:2.75,w:7.65,h:3.55,barDir:"col",showLegend:false,showTitle:false,
      showValue:true,dataLabelFormatCode:"0.00%",dataLabelPosition:"outEnd",dataLabelFormat:"0.00%",dataLabelFormatColor:C.forest,dataLabelFontSize:17,
      valAxisMinVal:0,valAxisMaxVal:1,valAxisMajorUnit:.25,valAxisLabelFormatCode:"0%",chartColors:[C.forest],
      catAxisLabelFontFace:fontFace,catAxisLabelFontSize:15,valAxisLabelFontSize:11,
      valGridLine:{color:"DFE5E0",width:1},showBorder:false
    });
    rect(slide,9,3.4,3.42,2.2,C.pale);
    tx(slide,"72.03% → 60.52%",9.2,3.65,3.02,.7,{fontSize:22,bold:true,align:"center",rtlMode:false,color:C.forest});
    tx(slide,"المصدر · جدول الحجوزات",9.2,4.62,3.02,.43,{fontSize:14,align:"center"});
  }else if(i===7){
    rect(slide,8.62,3.16,3.76,.11,C.mint);
    tx(slide,"لإدارة نون وفريق تنظيم ورشة Claude",1,3.7,11.4,.63,{fontSize:24,bold:true,color:C.forest});
    if(process.env.NOON_PRESENTER) tx(slide,process.env.NOON_PRESENTER,1,4.8,11.4,.42,{fontSize:19,color:C.muted});
  }else{
    if(i===4){
      rect(slide,5.05,2.9,3.24,.54,"31574B");
      tx(slide,"Dragon Valley Hub",5.12,2.94,3.10,.43,{align:"center",rtlMode:false,fontSize:17,color:C.ivory,bold:true});
      for(const x of [2.63,6.66,10.69]) path(slide,[[6.67,3.44],[6.67,3.64],[x,3.64],[x,3.98]],C.line,1.6);
    }
    const showMap=i===1 && process.env.NOON_FLOOR_MAP;
    const w=showMap?6.55:items.length===2?5.62:3.8,gap=.23,y=i===4?4:3.55;
    if(showMap){
      items.forEach((item,k)=>{
        const row=2.92+k*.78;
        rect(slide,5.87,row,w,.63,C.white);
        tx(slide,item,9.43,row+.08,2.76,.44,{fontSize:17,bold:true,color:C.forest});
        tx(slide,descriptions[i][k],6.02,row+.08,3.15,.44,{fontSize:13});
      });
    }else items.forEach((item,k)=>card(slide,item,descriptions[i][k],12.42-w-k*(w+gap),y,w,dark));
    if(i===1 && process.env.NOON_FLOOR_MAP){
      slide.addImage({path:process.env.NOON_FLOOR_MAP,x:.75,y:2.92,w:4.72,h:2.95,altText:"خريطة طوابق الفندق الخيالي من ملفات المشروع"});
      tx(slide,"خريطة الفندق الخيالي · من ملفات التحدي",.75,5.99,4.72,.33,{fontSize:11});
    }
    if(i===1) tx(slide,"250 ألف ريال افتراضية · خطة ثلاثة أشهر",showMap?5.87:.7,5.45,showMap?6.55:11.72,.5,{fontSize:17,color:C.muted});
    if(i===3) tx(slide,"مثال تعليمي · قارن الفترتين واذكر طريقة الحساب",.7,5.5,11.72,.5,{fontSize:17,color:soft});
    if(i===5) tx(slide,"المشكلة المكتشفة · البوابة تكرر وظيفة اختبار الغرفة",.7,5.5,11.72,.5,{fontSize:17,color:soft});
  }
  rect(slide,.72,6.89,11.72,.012,dark?C.line:"C9D7CE");
  rect(slide,.72,6.89,11.72*(i+1)/data.length,.025,dark?C.sand:C.forest);
  tx(slide,String(i+1).padStart(2,"0"),.72,6.99,.6,.24,{align:"left",rtlMode:false,fontSize:11,color:soft});
  tx(slide,"وادي التنانين  /  ورشة Claude",8.5,6.99,3.92,.25,{fontSize:11,color:soft});
  slide.addNotes(`المدة المستهدفة ${seconds} ثانية\n${notes}`);
}
deck.writeFile({fileName:process.argv[2]??"Noon_Dragon_Valley_v0.3.0.pptx"});
