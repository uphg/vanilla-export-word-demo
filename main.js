document.querySelector('#app').innerHTML = `
<div id="source-html">
  <h1>
    <center style="color: red;font-family: '微软雅黑'">Artificial Intelligence</center>
  </h1>
  <h2>Overview</h2>
  <p>
    Artificial Intelligence(AI) is an emerging technology demonstrating machine intelligence. The sub studies like <u><i>Neural Networks</i>, <i>Robatics</i> or <i>Machine Learning</i></u> are the parts of AI. This technology is expected to be a prime part of the real world in all levels.
  </p>
  <div>
    <img src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAFAAAABQCAMAAAC5zwKfAAAAPFBMVEX///8QX3X1+fnY6+oGWG/G3t8ATWYxo5P8/f0jnYzo8vKdy8mx2dUydYhYtKduvrNEq52Gwb54pbFWjp6S6lfVAAAD+0lEQVRYw62Y2YKDIAxFEQQSQAX9/3+dBPcW12leplPt4cINESPE/wPETwNASPghjCLE8AuVIwxAdYhRAvxEnx+EjGgtpjsSTwclZW3wxvTCE9CiujXpEySIDq2q6gSyRbT+2hcQSp0g6XL0EDR90NipG7kD0I43wlHqLRfoLnlJBBHQoi0jYSaC1tMn5c+JAM5yHCAhNQ44mjr/Famq9SmRbkE7I/UeCaAbtpe/CoEvgeiNGc6AvNJ2DsQ9kuZXV5XRQtLKidTx1EXTO7hwxNojJMk39aBcNQgXEVkkyL1PXwL9lveJpN8PgZwdEkCiK6eskefsVyC2evs7GCmyDfRJdzJ/qU92wTkSNhuJ54Oe/+uNv3TkExnG/BOO0gmkbhSDAtcbrhWNhmtHvpDAlcFiK8lt0+SsSZ0EuO9IAUkVoaU86eskYbFfDapAnffIBVJ5Vibd7A4lTl9Rcsu7jnwgYxgpuiEk1xyQjakrEx448sGMgW0IjaJHQMv7uTdVVVfuC5hL+h0iRp/nFzBS0tAPm9pUvXvmyKdKL/MSRS3IIW8GVajD9klgJF+Vl65tuTh+b8A7juyJeSPzjuYC+103bzqyJeZnvIydKmQ2eR/t04hSJi/cmz1yIDHJuimXr6eOzKGU064M7PAFj6pEagEOHpzvAktPvFeOzBrLAtNrgaXj4VtHcuaUBbY/FvjaEVs6vv7DEYu+KDC9F1hI6v84Ujxenz04L6Mo8L0jJYF0DHjtCJWGEvC9I9PR679legMsCfyHI2WB4R8CC3WLjjLvBRbrlvPtWyQevdaqRGfvXwkcz2MydBbxVwLHJz5N/SEyHr+Cwvj6px9NHf3Jewm/FD2depTHM5ZN0/MBnpnK35N5LJB4g6lNPTFJpr4l0x0A+UXL0GG2WpjsULhKzsO+Bc0y82pCbplwlZzHbQvm1RlZ73WeOVQqC9POo/XL6qq6oFMdJuehQN0PDSHMgs0yV2Y5OY8Fjiq1X7HVpHZlFhzC4w7DcjSGAnZhfjiE7UVXZT1wf2IX5s4h1NdtGu+DcgVsPc19WzmP6tb+OUDDxzZ5raQoqDUjc5w6XvULlyYDZmyXgnYLVo7YemaGDqO8FNjtuyoTdl2DEdtwDZHcZbsS6HEMu+fSGnTbNRBSBe+1vHQE6LaUuq5t2xhnjWvkpZ3WgF/db/Rw126MlM4ppbQO0yA0Styodfc6wjDGUbnMo+QxQtDqSfsWQGt603LlWT1vBOcuGe86srMZKPo+Jcp4zcPQONNIcJ+npj3MYbZB39MIPZsDTybcmAXESlkmiWSJd93YC+xNlee6MHKba53j6Bw8cETNOw7WVuEcb9v9OUukXLPoIekPRXxs+n14NQQAAAAASUVORK5CYII=" alt="logo">
  </div>
</div>
<div class="content-footer">
  <button id="btn">Export to word doc</button>
</div>
`

const btn = document.querySelector('#btn')

btn.addEventListener('click', exportWord)

function exportWord(){
  const header = `<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">
  <head>
    <meta charset="utf-8"><title>Export HTML to Word Document with JavaScript</title>
    <style>body{ color: blue; }</style>
  </head>
  <body>`
  const footer = '</body></html>'

  const sourceHTML = header+document.getElementById("source-html").innerHTML+footer;
  const source = 'data:application/vnd.ms-word;charset=utf-8,' + encodeURIComponent(sourceHTML);

  const fileDownload = document.createElement("a");
  document.body.appendChild(fileDownload);
  fileDownload.href = source;
  fileDownload.download = 'document.doc';
  fileDownload.click();
  document.body.removeChild(fileDownload);
}
