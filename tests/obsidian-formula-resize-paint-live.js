// Run through Obsidian CLI eval. Samples pixels between rapid resize updates.
(async()=>{
  const wait=ms=>new Promise(resolve=>setTimeout(resolve,ms));
  const assert=(value,message)=>{if(!value)throw Error(message);};
  const prior=app.workspace.activeLeaf,root=`Sheetian resize paint checks ${Date.now()}`;
  const cells=[];
  for(let r=0;r<2000;r++)for(let c=0;c<26;c++)cells.push({r,c,v:{v:r+c,bg:'#dceefe'}});
  const source=JSON.stringify([{id:'paint',name:'Resize paint',row:2000,column:26,celldata:cells}]);
  await app.vault.createFolder(root);
  const file=await app.vault.create(root+'/paint.sheet',source),leaf=app.workspace.getLeaf(true);
  let observer,frame;
  try{
    await leaf.openFile(file);app.workspace.setActiveLeaf(leaf,{focus:true});await wait(400);
    const view=leaf.view,canvas=view.contentEl.querySelector('.fortune-sheet-canvas');
    const placeholder=view.contentEl.querySelector('.fortune-sheet-canvas-placeholder');
    const overlay=view.contentEl.querySelector('.fortune-sheet-overlay');
    const bar=view.contentEl.querySelector('.fortune-fx-editor'),handle=view.contentEl.querySelector('.sheetian-formula-resize');
    const ctx=canvas.getContext('2d'),samples=[];
    const sample=source=>{
      if(!canvas.isConnected)return;
      const alphas=[0.25,0.5,0.9].map(fraction=>ctx.getImageData(Math.min(100,canvas.width-1),Math.floor((canvas.height-1)*fraction),1,1).data[3]);
      samples.push({source,bar:bar.offsetHeight,viewport:placeholder.clientHeight,overlay:overlay.clientHeight,canvas:canvas.clientHeight,alphas});
    };
    sample('initial');assert(samples[0].alphas.every(alpha=>alpha===255),'Paint fixture is not fully drawn');
    observer=new MutationObserver(()=>sample('mutation'));
    observer.observe(canvas,{attributes:true,attributeFilter:['width','height','style']});
    observer.observe(bar,{attributes:true,attributeFilter:['style']});
    const tick=()=>{sample('frame');frame=requestAnimationFrame(tick);};frame=requestAnimationFrame(tick);
    const capture=handle.setPointerCapture;handle.setPointerCapture=()=>{};
    const pointer=(type,y)=>handle.dispatchEvent(new PointerEvent(type,{pointerId:71,button:0,clientY:y,bubbles:true}));
    try{
      pointer('pointerdown',100);
      for(let n=0;n<48;n++)await new Promise(resolve=>requestAnimationFrame(()=>{
        // Several events in one frame, with a large reversal on the next.
        const height=n%2===0?350:65;
        for(let step=0;step<8;step++)pointer('pointermove',100+height-29-7+step);
        resolve();
      }));
      pointer('pointerup',136);
    }finally{handle.setPointerCapture=capture;}
    await wait(150);observer.disconnect();cancelAnimationFrame(frame);
    const blank=samples.filter(sample=>sample.alphas.some(alpha=>alpha!==255));
    const mismatched=samples.filter(sample=>Math.abs(sample.canvas-sample.viewport)>1||Math.abs(sample.overlay-sample.viewport)>1);
    const outcome={samples:samples.length,blankFrames:blank.filter(sample=>sample.source==='frame').length,blankUpdates:blank.length,mismatchedUpdates:mismatched.length,examples:[...blank,...mismatched].slice(0,5)};
    window.__sheetianResizePaintOutcome=outcome;
    assert(blank.length===0&&mismatched.length===0,'Rapid resize exposed an incomplete grid: '+JSON.stringify(outcome));
    assert(bar.offsetHeight===65,'The last rapid pointer position was lost');
    assert(view.getViewData()===source,'Rapid resizing changed workbook contents');
    return JSON.stringify({passed:3,...outcome,finalHeight:bar.offsetHeight});
  }finally{
    observer?.disconnect();if(frame!==undefined)cancelAnimationFrame(frame);
    await leaf.detach();await wait(100);if(prior)app.workspace.setActiveLeaf(prior,{focus:true});
    await app.vault.trash(file,true);await app.vault.trash(app.vault.getAbstractFileByPath(root),true);
  }
})()
