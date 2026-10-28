// RUA EZ-Fac: guided tour (first start, or the "?" button). No libraries.
// Each step: {tab, sel, titulo, texto, falta}. "tab" is the data-tour of a tab to open first, "sel" the data-tour
// of the element to highlight; if it is not on screen (e.g. the grid without a batch) the card shows in the middle
// with "falta" added. When the tour ends (or is skipped) it emits "ez_recorrido_fin" to the server.
(function () {
  const css = `
  .ezr-velo { position: fixed; inset: 0; z-index: 9000; }
  .ezr-foco { position: fixed; z-index: 9001; border-radius: 10px; pointer-events: none;
              box-shadow: 0 0 0 9999px rgba(13,13,82,.62), 0 0 0 3px #ff542c; transition: all .25s ease; }
  .ezr-tarjeta { position: fixed; z-index: 9002; width: 360px; max-width: calc(100vw - 32px); background: #fff;
                 border-radius: 14px; padding: 16px 18px 12px; box-shadow: 0 16px 40px rgba(0,0,0,.3);
                 max-height: calc(100vh - 24px); overflow-y: auto;
                 font-family: inherit; color: #1d2030; transition: top .25s ease, left .25s ease; }
  .ezr-tarjeta h3 { margin: 0 0 6px; font-size: 16px; color: #14147a; }
  .ezr-tarjeta p { margin: 0 0 8px; font-size: 14px; line-height: 1.4; }
  .ezr-tarjeta .ezr-falta { color: #6b7088; font-size: 12.5px; }
  .ezr-pie { display: flex; align-items: center; gap: 8px; margin-top: 10px; }
  .ezr-pie .ezr-n { color: #6b7088; font-size: 12px; margin-right: auto; }
  .ezr-pie button { border: 0; border-radius: 8px; padding: 6px 12px; font-size: 13px; cursor: pointer;
                    background: #eef0f8; color: #14147a; font-weight: 600; }
  .ezr-pie button.ezr-sig { background: #ff542c; color: #fff; }
  .ezr-saltar { position: absolute; top: 8px; right: 10px; border: 0; background: none; color: #6b7088;
                font-size: 12px; cursor: pointer; }`;

  let pasos = [], i = 0, velo, foco, tarjeta, elActual = null;

  function el(sel) {
    if (!sel) return null;
    const e = document.querySelector(`[data-tour="${sel}"]`);
    if (!e) return null;
    const r = e.getBoundingClientRect();
    return (r.width > 0 && r.height > 0) ? e : null;
  }

  function armar() {
    if (!document.getElementById("ezr-css")) {
      const s = document.createElement("style");
      s.id = "ezr-css"; s.textContent = css; document.head.appendChild(s);
    }
    velo = document.createElement("div"); velo.className = "ezr-velo";
    foco = document.createElement("div"); foco.className = "ezr-foco";
    tarjeta = document.createElement("div"); tarjeta.className = "ezr-tarjeta";
    document.body.append(velo, foco, tarjeta);
    window.addEventListener("resize", ubicar);
    window.addEventListener("scroll", ubicar, true);
    document.addEventListener("keydown", tecla);
  }

  function cerrar(terminado) {
    [velo, foco, tarjeta].forEach(x => x && x.remove());
    window.removeEventListener("resize", ubicar);
    window.removeEventListener("scroll", ubicar, true);
    document.removeEventListener("keydown", tecla);
    velo = foco = tarjeta = null;
    if (window.emitEvent) emitEvent("ez_recorrido_fin", {terminado: !!terminado});
  }

  function tecla(ev) {
    if (ev.key === "Escape") cerrar(false);
    else if (ev.key === "ArrowRight" || ev.key === "Enter") ir(i + 1);
    else if (ev.key === "ArrowLeft") ir(i - 1);
  }

  function ubicar() {
    if (!tarjeta) return;
    const e = elActual && document.body.contains(elActual) ? elActual : null;
    tarjeta.style.width = "360px";                       // a previous step may have narrowed it
    const tw = tarjeta.offsetWidth, th = tarjeta.offsetHeight, vw = innerWidth, vh = innerHeight, m = 12;
    if (!e) {
      foco.style.cssText = "top:50%;left:50%;width:0;height:0;";
      tarjeta.style.top = Math.max(m, (vh - th) / 2) + "px";
      tarjeta.style.left = Math.max(m, (vw - tw) / 2) + "px";
      return;
    }
    const r = e.getBoundingClientRect(), p = 6, ANCHO = 360, MINIMO = 260;
    foco.style.cssText = `top:${r.top - p}px;left:${r.left - p}px;width:${r.width + 2 * p}px;height:${r.height + 2 * p}px;`;
    const fija = (v, min, max) => Math.max(min, Math.min(v, max));   // never outside the window
    const lugar = (top, left) => {
      tarjeta.style.top = fija(top, m, vh - tarjeta.offsetHeight - m) + "px";
      tarjeta.style.left = fija(left, m, vw - tarjeta.offsetWidth - m) + "px";
    };
    tarjeta.style.width = ANCHO + "px";
    let th2 = tarjeta.offsetHeight;
    if (r.bottom + p + m + th2 <= vh - m) return lugar(r.bottom + p + m, r.left);      // below the element
    if (r.top - p - m - th2 >= m) return lugar(r.top - p - m - th2, r.left);           // above
    const derecha = vw - r.right - p - 2 * m, izquierda = r.left - p - 2 * m;        // tall sections: beside,
    const lado = Math.max(derecha, izquierda);                                         // narrower if needed
    if (lado >= MINIMO) {
      tarjeta.style.width = Math.min(ANCHO, lado) + "px";
      const w = tarjeta.offsetWidth;
      return lugar(r.top, derecha >= izquierda ? r.right + p + m : r.left - p - m - w);
    }
    lugar(vh, vw);                                             // no room anywhere: bottom right corner
  }

  function mostrar() {
    const s = pasos[i];
    elActual = el(s.sel);
    if (elActual) elActual.scrollIntoView({block: "center", inline: "nearest", behavior: "instant"});
    const falta = (!elActual && s.sel && s.falta) ? `<p class="ezr-falta">${s.falta}</p>` : "";
    tarjeta.innerHTML = `<button class="ezr-saltar" title="Salir del recorrido (Esc)">Saltar ✕</button>
      <h3>${s.titulo}</h3><p>${s.texto}</p>${falta}
      <div class="ezr-pie"><span class="ezr-n">${i + 1} de ${pasos.length}</span>
      ${i > 0 ? '<button class="ezr-ant">Anterior</button>' : ""}
      <button class="ezr-sig">${i === pasos.length - 1 ? "Terminar" : "Siguiente"}</button></div>`;
    tarjeta.querySelector(".ezr-saltar").onclick = () => cerrar(false);
    tarjeta.querySelector(".ezr-sig").onclick = () => ir(i + 1);
    const ant = tarjeta.querySelector(".ezr-ant");
    if (ant) ant.onclick = () => ir(i - 1);
    requestAnimationFrame(ubicar);
  }

  function ir(n) {
    if (n >= pasos.length) return cerrar(true);
    if (n < 0) return;
    i = n;
    const s = pasos[i];
    const tab = s.tab && document.querySelector(`[data-tour="${s.tab}"]`);
    if (tab && !tab.classList.contains("q-tab--active")) {
      tab.click();
      setTimeout(mostrar, 450);                        // the tab panel needs a moment to render
    } else {
      mostrar();
    }
  }

  window.ezRecorrido = function (lista) {
    if (velo) cerrar(false);
    pasos = lista || [];
    if (!pasos.length) return;
    armar();
    ir(0);
  };
})();
