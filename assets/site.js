// matiasdiaz.cl · JS compartido
(function(){
  // Reveal on scroll
  var io=new IntersectionObserver(function(es){es.forEach(function(e){if(e.isIntersecting){e.target.classList.add('in');io.unobserve(e.target)}})},{threshold:.12});
  document.querySelectorAll('.reveal').forEach(function(el){io.observe(el)});

  // Onda del hero (solo si existe #wave)
  var cv=document.getElementById('wave');
  if(cv){
    var ctx=cv.getContext('2d');
    var reduced=matchMedia('(prefers-reduced-motion: reduce)').matches;
    function size(){cv.width=cv.offsetWidth;cv.height=cv.offsetHeight}
    size();addEventListener('resize',size);
    var t=0;
    var lines=[{c:'rgba(67,224,195,.35)',a:.9,f:1},{c:'rgba(61,143,214,.25)',a:.6,f:1.6},{c:'rgba(67,224,195,.12)',a:1.3,f:.7}];
    (function draw(){
      ctx.clearRect(0,0,cv.width,cv.height);
      lines.forEach(function(L){
        ctx.beginPath();
        for(var x=0;x<=cv.width;x+=6){
          var y=cv.height*.62+Math.sin((x*.006*L.f)+t*L.f)*38*L.a+Math.sin((x*.017)+t*1.7)*8;
          x===0?ctx.moveTo(x,y):ctx.lineTo(x,y);
        }
        ctx.strokeStyle=L.c;ctx.lineWidth=1.6;ctx.stroke();
      });
      t+=0.012;
      if(!reduced)requestAnimationFrame(draw);
    })();
  }

  // YouTube: click-to-play (no carga iframes hasta el clic)
  document.querySelectorAll('.yt[data-id]').forEach(function(el){
    var id=el.getAttribute('data-id');
    var img=document.createElement('img');
    img.loading='lazy';img.alt='';
    img.src='https://i.ytimg.com/vi/'+id+'/hqdefault.jpg';
    var play=document.createElement('div');play.className='play';play.innerHTML='<i>▶</i>';
    el.appendChild(img);el.appendChild(play);
    el.setAttribute('role','button');el.setAttribute('tabindex','0');
    el.setAttribute('aria-label','Reproducir video');
    function go(){
      var f=document.createElement('iframe');
      f.src='https://www.youtube-nocookie.com/embed/'+id+'?autoplay=1';
      f.allow='accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture';
      f.allowFullscreen=true;
      el.innerHTML='';el.appendChild(f);
    }
    el.addEventListener('click',go);
    el.addEventListener('keydown',function(e){if(e.key==='Enter'||e.key===' '){e.preventDefault();go()}});
  });
})();
