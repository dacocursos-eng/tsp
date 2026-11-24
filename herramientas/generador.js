/* Generador de evidencias TSP: lee un proyecto de GitLab (o el CSV de elementos de trabajo)
   y arma un libro de Excel con hojas de datos y tablas con fórmulas. */
(function (root) {
  "use strict";

  var FASES = ["Requisitos", "Diseño", "Revisión de diseño", "Implementación", "Revisión de código",
    "Inspección", "Pruebas", "Integración", "Corrección", "Gestión"];

  var TIPOS = ["10 documentación", "20 sintaxis", "30 construcción", "40 asignación", "50 interfaz",
    "60 chequeo", "70 datos", "80 función", "90 sistema", "100 entorno"];

  function norm(s) {
    return String(s == null ? "" : s).normalize("NFD").replace(/[̀-ͯ]/g, "")
      .toLowerCase().replace(/[^a-z0-9: ]+/g, " ").replace(/\s+/g, " ").trim();
  }

  // Variantes aceptadas para cada fase, normalizadas. Se buscan de la más larga a la más corta.
  var ALIAS = [
    ["revision de diseno", "Revisión de diseño"], ["revision del diseno", "Revisión de diseño"],
    ["revision de codigo", "Revisión de código"], ["revision del codigo", "Revisión de código"],
    ["requisitos", "Requisitos"], ["requisito", "Requisitos"], ["requerimientos", "Requisitos"],
    ["diseno", "Diseño"], ["implementacion", "Implementación"], ["codificacion", "Implementación"],
    ["programacion", "Implementación"], ["inspeccion", "Inspección"], ["pruebas", "Pruebas"],
    ["prueba", "Pruebas"], ["integracion", "Integración"], ["correccion", "Corrección"],
    ["gestion", "Gestión"], ["reunion", "Gestión"], ["reuniones", "Gestión"], ["planeacion", "Gestión"],
    ["planificacion", "Gestión"], ["testing", "Pruebas"]
  ].sort(function (a, b) { return b[0].length - a[0].length; });

  function fase(texto) {
    var n = norm(texto);
    if (!n) return "";
    for (var i = 0; i < ALIAS.length; i++) {
      var a = ALIAS[i][0];
      if (n === a || (" " + n + " ").indexOf(" " + a + " ") >= 0) return ALIAS[i][1];
    }
    return "";
  }

  function tipoDefecto(labels) {
    for (var i = 0; i < labels.length; i++) {
      var n = norm(labels[i]);
      if (n.indexOf("tipo") === 0) {
        var m = n.match(/(\d{2,3})/);
        if (m) {
          for (var j = 0; j < TIPOS.length; j++) if (TIPOS[j].split(" ")[0] === m[1]) return TIPOS[j];
        }
      }
    }
    return "";
  }

  function faseEtiqueta(labels, prefijo) {
    for (var i = 0; i < labels.length; i++) {
      var n = norm(labels[i]);
      if (n.indexOf(prefijo) === 0) return fase(n.slice(prefijo.length).replace(/^[: ]+/, ""));
    }
    return "";
  }

  // Impacto en masculino (alto, medio, bajo) y probabilidad en femenino (alta, media, baja).
  function nivel(labels, prefijo) {
    var fem = prefijo === "probabilidad";
    for (var i = 0; i < labels.length; i++) {
      var n = norm(labels[i]);
      if (n.indexOf(prefijo) === 0) {
        var v = n.slice(prefijo.length).replace(/^[: ]+/, "");
        if (v.indexOf("alt") === 0) return fem ? "Alta" : "Alto";
        if (v.indexOf("med") === 0) return fem ? "Media" : "Medio";
        if (v.indexOf("baj") === 0) return fem ? "Baja" : "Bajo";
      }
    }
    return "";
  }

  function clase(labels, tipo) {
    var ns = labels.map(norm);
    if (ns.indexOf("defecto") >= 0) return "Defecto";
    if (ns.indexOf("riesgo") >= 0) return "Riesgo";
    if (ns.indexOf("reunion semanal") >= 0) return "Reunión semanal";
    if (ns.indexOf("relanzamiento") >= 0) return "Relanzamiento";
    if (ns.indexOf("requisito") >= 0) return "Requisito";
    if (ns.indexOf("tarea") >= 0) return "Tarea";
    if (norm(tipo) === "task") return "Tarea";
    return "Otro";
  }

  // Fecha local de Colombia como fecha de Excel (medianoche UTC del mismo día).
  function fechaLocal(v) {
    if (!v) return null;
    var s = String(v);
    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
      var p = s.split("-");
      return new Date(Date.UTC(+p[0], +p[1] - 1, +p[2]));
    }
    var d = new Date(s);
    if (isNaN(d)) return null;
    var t = d.toLocaleDateString("en-CA", { timeZone: "America/Bogota" });
    var q = t.split("-");
    return new Date(Date.UTC(+q[0], +q[1] - 1, +q[2]));
  }

  // Duración de GitLab a horas: segundos o texto como "1w 2d 3h 30m".
  function horas(v) {
    if (v == null || v === "") return 0;
    if (typeof v === "number") return v / 3600;
    var s = String(v).trim();
    if (/^\d+(\.\d+)?$/.test(s)) return parseFloat(s) / 3600;
    var total = 0, re = /(\d+(?:[.,]\d+)?)\s*(mo|w|d|h|m|s)/gi, m;
    var f = { mo: 160, w: 40, d: 8, h: 1, m: 1 / 60, s: 1 / 3600 };
    while ((m = re.exec(s))) total += parseFloat(m[1].replace(",", ".")) * f[m[2].toLowerCase()];
    return total;
  }

  function redondear(x) { return Math.round(x * 100) / 100; }

  // Une una lista en español: "a", "a y b", "a, b y c". Si se indica, agrega al final "y otras" u "y otros".
  function unir(xs, resto) {
    xs = xs.slice();
    if (resto) xs.push(resto);
    if (xs.length <= 1) return xs.join("");
    return xs.slice(0, -1).join(", ") + " y " + xs[xs.length - 1];
  }

  // Número con el sustantivo y el verbo en singular o en plural.
  function cuenta(n, singular, plural) { return n + " " + (n === 1 ? singular : plural); }

  function cierres(texto) {
    var out = [], re = /\b(?:close[sd]?|closing|fix(?:e[sd])?|fixing|resolve[sd]?|resolving|implement(?:s|ed)?|implementing)\b:?\s+((?:(?:#\d+)(?:\s*(?:,|and|y)\s*)?)+)/gi, m;
    while ((m = re.exec(texto || ""))) {
      (m[1].match(/#(\d+)/g) || []).forEach(function (x) { out.push(+x.slice(1)); });
    }
    return out;
  }

  /* ---------- Lectura de GitLab ---------- */

  function rutaProyecto(url) {
    var s = String(url || "").trim().replace(/[\s.,;:)"']+$/, "").replace(/\.git$/, "");
    s = s.replace(/^https?:\/\/[^/]+\//, "").replace(/\/-\/.*$/, "").replace(/\/+$/, "");
    return s;
  }

  function Cliente(token, base, avisar) {
    base = base || "https://gitlab.com";
    function get(ruta) {
      return fetch(base + "/api/v4" + ruta, { headers: { "PRIVATE-TOKEN": token } }).then(function (r) {
        if (r.status === 401) throw new Error("GitLab no aceptó el token. Revisen que esté copiado completo, que no haya vencido y que tenga el alcance read_api.");
        if (r.status === 404) throw new Error("GitLab no encontró el proyecto o no tienen acceso a él con ese token. Revisen la dirección del proyecto.");
        if (!r.ok) throw new Error("GitLab respondió con el código " + r.status + " al consultar " + ruta);
        return r.json();
      });
    }
    function todo(ruta) {
      var out = [], pag = 1, sep = ruta.indexOf("?") >= 0 ? "&" : "?";
      function sig() {
        return get(ruta + sep + "per_page=100&page=" + pag).then(function (lote) {
          out = out.concat(lote);
          if (lote.length === 100 && pag < 50) { pag++; return sig(); }
          return out;
        });
      }
      return sig();
    }
    function gql(query, variables) {
      return fetch(base + "/api/graphql", {
        method: "POST",
        headers: { "PRIVATE-TOKEN": token, "Content-Type": "application/json" },
        body: JSON.stringify({ query: query, variables: variables || {} })
      }).then(function (r) {
        if (!r.ok) throw new Error("GraphQL respondió con el código " + r.status);
        return r.json();
      }).then(function (j) {
        if (j.errors && j.errors.length) throw new Error(j.errors[0].message);
        return j.data;
      });
    }
    return { get: get, todo: todo, gql: gql };
  }

  function leerGitLab(url, token, progreso) {
    progreso = progreso || function () {};
    var ruta = rutaProyecto(url);
    if (!ruta || ruta.indexOf("/") < 0) return Promise.reject(new Error("La dirección del proyecto no es válida. Copien la dirección completa de la página principal del proyecto."));
    if (!token) return Promise.reject(new Error("Falta el token."));
    var c = Cliente(token);
    var datos = { modo: "gitlab", proyecto: {}, items: [], tiempo: [], fusiones: [], hitos: [], avisos: [] };
    var pid;
    progreso("Consultando el proyecto");
    return c.get("/projects/" + encodeURIComponent(ruta)).then(function (p) {
      pid = p.id;
      datos.proyecto = { nombre: p.name_with_namespace, url: p.web_url, ruta: p.path_with_namespace };
      progreso("Leyendo hitos");
      return c.todo("/projects/" + pid + "/milestones?include_ancestors=true").catch(function () {
        return c.todo("/projects/" + pid + "/milestones");
      });
    }).then(function (ms) {
      ms.forEach(function (m) { agregarHito(datos, m.title, m.start_date, m.due_date); });
      progreso("Leyendo elementos de trabajo");
      return c.todo("/projects/" + pid + "/issues?state=all&scope=all");
    }).then(function (issues) {
      issues.forEach(function (i) {
        var labels = i.labels || [];
        if (i.milestone) agregarHito(datos, i.milestone.title, i.milestone.start_date, i.milestone.due_date);
        datos.items.push({
          iid: i.iid, tipo: i.issue_type === "task" ? "Task" : (i.issue_type === "issue" ? "Issue" : (i.issue_type || "Issue")),
          titulo: i.title, estado: i.state === "closed" ? "Cerrado" : "Abierto", etiquetas: labels,
          padreIid: null, padreTitulo: "",
          responsable: (i.assignees && i.assignees[0] ? i.assignees[0].name : (i.assignee ? i.assignee.name : "")),
          hito: i.milestone ? i.milestone.title : "",
          estimado: horas((i.time_stats || {}).time_estimate || 0),
          dedicado: horas((i.time_stats || {}).total_time_spent || 0),
          creado: fechaLocal(i.created_at), cerrado: fechaLocal(i.closed_at)
        });
      });
      progreso("Leyendo la jerarquía de requisitos y tareas");
      return leerPadres(c, ruta).then(function (padres) {
        datos.items.forEach(function (it) {
          var p = padres[it.iid];
          if (p) { it.padreIid = p.iid; it.padreTitulo = p.title; }
        });
      }).catch(function () {
        datos.avisos.push("No fue posible leer a qué requisito pertenece cada tarea. La columna Requisito queda vacía.");
      });
    }).then(function () {
      progreso("Leyendo los registros de tiempo");
      return leerTiempo(c, ruta);
    }).then(function (tl) {
      var porIid = {};
      datos.items.forEach(function (it) { porIid[it.iid] = it; });
      tl.forEach(function (t) {
        var texto = t.summary || (t.note && t.note.body) || "";
        var it = t.issue ? porIid[+t.issue.iid] : null;
        datos.tiempo.push({
          fecha: fechaLocal(t.spentAt), persona: t.user ? (t.user.name || t.user.username) : "",
          usuario: t.user ? t.user.username : "",
          elemento: t.issue ? "#" + t.issue.iid : (t.mergeRequest ? "!" + t.mergeRequest.iid : ""),
          iid: t.issue ? +t.issue.iid : null, mr: t.mergeRequest ? +t.mergeRequest.iid : null,
          titulo: t.issue ? t.issue.title : (t.mergeRequest ? t.mergeRequest.title : ""),
          clase: it ? clase(it.etiquetas, it.tipo) : (t.mergeRequest ? "Solicitud de fusión" : ""),
          horas: redondear(t.timeSpent / 3600), texto: texto, fase: fase(texto)
        });
      });
      progreso("Leyendo las solicitudes de fusión");
      return c.todo("/projects/" + pid + "/merge_requests?state=all&scope=all");
    }).then(function (mrs) {
      var cadena = Promise.resolve();
      mrs.forEach(function (m, k) {
        cadena = cadena.then(function () {
          progreso("Leyendo la solicitud de fusión " + (k + 1) + " de " + mrs.length);
          return leerFusion(c, pid, m).then(function (f) { datos.fusiones.push(f); });
        });
      });
      return cadena;
    }).then(function () {
      completar(datos);
      return datos;
    });
  }

  function leerPadres(c, ruta) {
    var padres = {};
    var q = "query($p:ID!,$a:String){project(fullPath:$p){workItems(first:50,after:$a){pageInfo{hasNextPage endCursor} nodes{iid widgets{... on WorkItemWidgetHierarchy{parent{iid title}}}}}}}";
    function pag(after, n) {
      return c.gql(q, { p: ruta, a: after }).then(function (d) {
        var w = d.project.workItems;
        w.nodes.forEach(function (n) {
          (n.widgets || []).forEach(function (x) { if (x && x.parent) padres[+n.iid] = { iid: +x.parent.iid, title: x.parent.title }; });
        });
        if (w.pageInfo.hasNextPage && n < 40) return pag(w.pageInfo.endCursor, n + 1);
        return padres;
      });
    }
    return pag(null, 0);
  }

  function leerTiempo(c, ruta) {
    var out = [];
    var q = "query($p:ID!,$a:String){project(fullPath:$p){timelogs(first:100,after:$a){pageInfo{hasNextPage endCursor} nodes{timeSpent spentAt summary user{username name} issue{iid title} mergeRequest{iid title} note{body}}}}}";
    function pag(after, n) {
      return c.gql(q, { p: ruta, a: after }).then(function (d) {
        var t = d.project.timelogs;
        out = out.concat(t.nodes);
        if (t.pageInfo.hasNextPage && n < 100) return pag(t.pageInfo.endCursor, n + 1);
        return out;
      });
    }
    return pag(null, 0);
  }

  function leerFusion(c, pid, m) {
    var base = "/projects/" + pid + "/merge_requests/" + m.iid;
    var f = {
      iid: m.iid, titulo: m.title, autor: m.author ? m.author.name : "",
      estado: { merged: "Fusionada", opened: "Abierta", closed: "Cerrada", locked: "Bloqueada" }[m.state] || m.state,
      tareas: cierres(m.description), rama: m.source_branch,
      revisor: (m.reviewers || []).map(function (r) { return r.name; }).join(", "),
      aprobadaPor: "", fusionada: fechaLocal(m.merged_at), creada: fechaLocal(m.created_at),
      agregadas: 0, eliminadas: 0, comentarios: 0
    };
    var diffs = c.todo(base + "/diffs").catch(function () {
      return c.get(base + "/changes").then(function (x) { return x.changes || []; });
    }).then(function (ds) {
      ds.forEach(function (d) {
        String(d.diff || "").split("\n").forEach(function (l) {
          if (l.indexOf("+++") === 0 || l.indexOf("---") === 0) return;
          if (l.charAt(0) === "+") f.agregadas++;
          else if (l.charAt(0) === "-") f.eliminadas++;
        });
      });
    }).catch(function () {});
    var notas = c.todo(base + "/notes").then(function (ns) {
      f.comentarios = ns.filter(function (n) { return !n.system && n.type === "DiffNote"; }).length;
    }).catch(function () {});
    var apr = c.get(base + "/approvals").then(function (a) {
      f.aprobadaPor = (a.approved_by || []).map(function (x) { return x.user.name; }).join(", ");
    }).catch(function () {});
    var cl = f.tareas.length ? Promise.resolve() : c.get(base + "/closes_issues").then(function (xs) {
      f.tareas = xs.map(function (x) { return x.iid; });
    }).catch(function () {});
    return Promise.all([diffs, notas, apr, cl]).then(function () { return f; });
  }

  function agregarHito(datos, titulo, inicio, fin) {
    if (!titulo) return;
    var h = datos.hitos.filter(function (x) { return x.titulo === titulo; })[0];
    if (!h) { h = { titulo: titulo, inicio: null, fin: null }; datos.hitos.push(h); }
    if (inicio && !h.inicio) h.inicio = fechaLocal(inicio);
    if (fin && !h.fin) h.fin = fechaLocal(fin);
  }

  /* ---------- Lectura del CSV ---------- */

  function parseCSV(texto) {
    texto = String(texto).replace(/^﻿/, "");
    var filas = [], fila = [], campo = "", q = false, i, ch;
    for (i = 0; i < texto.length; i++) {
      ch = texto[i];
      if (q) {
        if (ch === '"') { if (texto[i + 1] === '"') { campo += '"'; i++; } else q = false; }
        else campo += ch;
      } else if (ch === '"') q = true;
      else if (ch === ",") { fila.push(campo); campo = ""; }
      else if (ch === "\n" || ch === "\r") {
        if (ch === "\r" && texto[i + 1] === "\n") i++;
        fila.push(campo); campo = ""; filas.push(fila); fila = [];
      } else campo += ch;
    }
    if (campo !== "" || fila.length) { fila.push(campo); filas.push(fila); }
    return filas.filter(function (f) { return f.some(function (x) { return x !== ""; }); });
  }

  function leerCSV(texto, nombre) {
    var filas = parseCSV(texto);
    if (filas.length < 2) throw new Error("El archivo no tiene filas de datos.");
    var enc = filas[0].map(function (h) { return norm(h); });
    function col() {
      for (var i = 0; i < arguments.length; i++) {
        var k = enc.indexOf(norm(arguments[i]));
        if (k >= 0) return k;
      }
      return -1;
    }
    var C = {
      iid: col("IID", "Issue ID", "Work item IID", "ID"), titulo: col("Title", "Título"),
      tipo: col("Type", "Work item type", "Tipo"), estado: col("State", "Estado"),
      resp: col("Assignee", "Assignees", "Persona asignada"), hito: col("Milestone", "Hito"),
      labels: col("Labels", "Etiquetas"), est: col("Time Estimate", "Estimate"),
      ded: col("Time Spent", "Spent"), creado: col("Created At (UTC)", "Created At"),
      cerrado: col("Closed At (UTC)", "Closed At"), inicio: col("Start Date", "Start date"),
      vence: col("Due Date", "Due date"), piid: col("Parent IID"), ptit: col("Parent Title")
    };
    if (C.titulo < 0 || C.iid < 0) throw new Error("El archivo no parece ser el CSV de elementos de trabajo de GitLab: no tiene las columnas de título y número.");
    var datos = { modo: "csv", proyecto: { nombre: nombre || "", url: "", ruta: "" }, items: [], tiempo: [], fusiones: [], hitos: [], avisos: [] };
    var rangos = {};
    function v(f, k) { return k >= 0 ? (f[k] || "").trim() : ""; }
    filas.slice(1).forEach(function (f) {
      var labels = v(f, C.labels) ? v(f, C.labels).split(",").map(function (x) { return x.trim(); }).filter(Boolean) : [];
      var hito = v(f, C.hito);
      var it = {
        iid: parseInt(v(f, C.iid), 10), tipo: v(f, C.tipo) || "Issue", titulo: v(f, C.titulo),
        estado: /closed|cerrad/i.test(v(f, C.estado)) ? "Cerrado" : "Abierto", etiquetas: labels,
        padreIid: v(f, C.piid) ? parseInt(v(f, C.piid), 10) : null, padreTitulo: v(f, C.ptit),
        responsable: v(f, C.resp).split(",")[0].trim(), hito: hito,
        estimado: redondear(horas(v(f, C.est))), dedicado: redondear(horas(v(f, C.ded))),
        creado: fechaLocal(v(f, C.creado).replace(" UTC", "Z").replace(" ", "T")),
        cerrado: fechaLocal(v(f, C.cerrado).replace(" UTC", "Z").replace(" ", "T"))
      };
      if (isNaN(it.iid)) return;
      datos.items.push(it);
      if (hito) {
        var a = fechaLocal(v(f, C.inicio)), b = fechaLocal(v(f, C.vence));
        var r = rangos[hito] || (rangos[hito] = { ini: null, fin: null });
        [a, b].forEach(function (d) {
          if (!d) return;
          if (!r.ini || d < r.ini) r.ini = d;
          if (!r.fin || d > r.fin) r.fin = d;
        });
      }
    });
    Object.keys(rangos).forEach(function (h) { datos.hitos.push({ titulo: h, inicio: rangos[h].ini, fin: rangos[h].fin }); });
    datos.avisos.push("El CSV no trae el detalle de los registros de tiempo ni las solicitudes de fusión. Diligencien las hojas Tiempo y Fusiones a partir del informe de tiempo del grupo y de cada solicitud de fusión.");
    datos.avisos.push("Las fechas de cada semana se calcularon con las fechas de las tareas de cada hito. Revísenlas en la hoja Semanas y corríjanlas con las fechas de sus hitos si hace falta.");
    completar(datos);
    return datos;
  }

  /* ---------- Revisión de los datos ---------- */

  function completar(d) {
    d.hitos.sort(function (a, b) {
      var x = a.inicio ? a.inicio.getTime() : Infinity, y = b.inicio ? b.inicio.getTime() : Infinity;
      return x - y || a.titulo.localeCompare(b.titulo, "es", { numeric: true });
    });
    d.items.sort(function (a, b) { return a.iid - b.iid; });
    d.items.forEach(function (it) {
      it.clase = clase(it.etiquetas, it.tipo);
      it.estimado = redondear(it.estimado);
      it.dedicado = redondear(it.dedicado);
    });
    var porIid = {};
    d.items.forEach(function (it) { porIid[it.iid] = it; });
    d.items.forEach(function (it) {
      if (it.padreIid && !it.padreTitulo && porIid[it.padreIid]) it.padreTitulo = porIid[it.padreIid].titulo;
    });
    d.defectos = d.items.filter(function (it) { return it.clase === "Defecto"; }).map(function (it) {
      return { it: it, tipo: tipoDefecto(it.etiquetas), inyectado: faseEtiqueta(it.etiquetas, "inyectado"), removido: faseEtiqueta(it.etiquetas, "removido") };
    });
    d.riesgos = d.items.filter(function (it) { return it.clase === "Riesgo"; }).map(function (it) {
      return { it: it, impacto: nivel(it.etiquetas, "impacto"), probabilidad: nivel(it.etiquetas, "probabilidad") };
    });
    // Inspección: horas con la fase Inspección en la tarea que cierra cada solicitud o en la solicitud misma.
    d.fusiones.forEach(function (f) {
      var personas = {}, h = 0;
      d.tiempo.forEach(function (t) {
        if (t.fase !== "Inspección") return;
        if ((t.iid && f.tareas.indexOf(t.iid) >= 0) || (t.mr && t.mr === f.iid)) { h += t.horas; personas[t.persona] = 1; }
      });
      f.horasInspeccion = redondear(h);
      f.inspectores = Object.keys(personas).length;
      f.tarea = f.tareas.length ? f.tareas.map(function (x) { return "#" + x; }).join(", ") : "";
    });
    d.fusiones.sort(function (a, b) { return a.iid - b.iid; });
    d.tiempo.sort(function (a, b) { return (a.fecha || 0) - (b.fecha || 0); });
    d.personas = [];
    function per(p) { if (p && d.personas.indexOf(p) < 0) d.personas.push(p); }
    d.tiempo.forEach(function (t) { per(t.persona); });
    d.items.forEach(function (it) { if (it.clase === "Tarea") per(it.responsable); });
    d.personas.sort(function (a, b) { return a.localeCompare(b, "es"); });

    var av = d.avisos;
    var sinFase = d.tiempo.filter(function (t) { return !t.fase; });
    if (sinFase.length) {
      var ej = sinFase.slice(0, 5).map(function (t) { return "\"" + (t.texto || "vacío") + "\" (" + t.persona + ", " + t.elemento + ")"; });
      ej = ej.length > 1 ? ej.slice(0, -1).join("; ") + " y " + ej[ej.length - 1] : ej.join("");
      av.push(sinFase.length === 1 ? "1 registro de tiempo no tiene una fase reconocible en el resumen: " + ej + "."
        : sinFase.length + " registros de tiempo no tienen una fase reconocible en el resumen. " + (sinFase.length > 5 ? "Algunos ejemplos: " : "Son estos: ") + ej + ".");
    }
    var tareas = d.items.filter(function (it) { return it.clase === "Tarea" && it.hito; });
    function lista(xs, otros) { return unir(xs.slice(0, 8).map(function (it) { return "#" + it.iid; }), xs.length > 8 ? otros : ""); }
    var sinEst = tareas.filter(function (it) { return !it.estimado; });
    if (sinEst.length) av.push(cuenta(sinEst.length, "tarea no tiene", "tareas no tienen") + " estimación: " + lista(sinEst, "otras") + ".");
    var sinHito = d.items.filter(function (it) { return it.clase === "Tarea" && !it.hito && it.estado !== "Cerrado"; });
    if (sinHito.length) av.push(cuenta(sinHito.length, "tarea abierta no tiene", "tareas abiertas no tienen") + " hito: " + lista(sinHito, "otras") + ".");
    var sinResp = tareas.filter(function (it) { return !it.responsable; });
    if (sinResp.length) av.push(cuenta(sinResp.length, "tarea no tiene", "tareas no tienen") + " persona asignada: " + lista(sinResp, "otras") + ".");
    var defInc = d.defectos.filter(function (x) { return !x.tipo || !x.inyectado || !x.removido; });
    if (defInc.length) av.push(cuenta(defInc.length, "defecto no tiene", "defectos no tienen") + " completas las etiquetas de tipo, fase inyectada y fase removida: " + lista(defInc.map(function (x) { return x.it; }), "otros") + ".");
    var riInc = d.riesgos.filter(function (x) { return !x.impacto || !x.probabilidad; });
    if (riInc.length) av.push(cuenta(riInc.length, "riesgo no tiene", "riesgos no tienen") + " etiqueta de impacto o de probabilidad: " + lista(riInc.map(function (x) { return x.it; }), "otros") + ".");
    var mrSin = d.fusiones.filter(function (f) { return !f.tareas.length; });
    if (mrSin.length) av.push(cuenta(mrSin.length, "solicitud de fusión no indica la tarea que cierra", "solicitudes de fusión no indican la tarea que cierran") + " con Closes #: " + unir(mrSin.slice(0, 8).map(function (f) { return "!" + f.iid; }), mrSin.length > 8 ? "otras" : "") + ".");
    var hSin = d.hitos.filter(function (h) { return !h.inicio; });
    if (hSin.length) av.push((hSin.length === 1 ? "El hito " : "Los hitos ") + unir(hSin.map(function (h) { return h.titulo; })) + (hSin.length === 1 ? " no tiene" : " no tienen") + " fecha de inicio. Escríbanla en la hoja Semanas para que el tiempo se asigne a la semana correcta.");
    return d;
  }

  function resumen(d) {
    var tareas = d.items.filter(function (it) { return it.clase === "Tarea" && it.hito; });
    var tot = tareas.reduce(function (s, it) { return s + it.estimado; }, 0);
    var ev = tareas.filter(function (it) { return it.estado === "Cerrado"; }).reduce(function (s, it) { return s + it.estimado; }, 0);
    var lineas = d.fusiones.reduce(function (s, f) { return s + f.agregadas; }, 0);
    return {
      elementos: d.items.length,
      tareas: tareas.length,
      terminadas: tareas.filter(function (it) { return it.estado === "Cerrado"; }).length,
      horas: redondear(d.tiempo.reduce(function (s, t) { return s + t.horas; }, 0)),
      registros: d.tiempo.length,
      valorGanado: tot ? redondear(ev / tot * 100) : 0,
      fusiones: d.fusiones.length,
      lineas: lineas,
      defectos: d.defectos.length,
      riesgos: d.riesgos.length,
      personas: d.personas.length
    };
  }

  /* ---------- Construcción del Excel ---------- */

  var AZUL = "FF1F4E79", CLARO = "FFE7EEF6", GRIS = "FFF1EFE9";

  function encabezado(ws, fila) {
    ws.getRow(fila).eachCell(function (c) {
      c.font = { bold: true, color: { argb: "FFFFFFFF" } };
      c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: AZUL } };
      c.alignment = { vertical: "middle", wrapText: true };
    });
    ws.getRow(fila).height = 30;
  }

  function hojaDatos(wb, nombre, cols, filas) {
    var ws = wb.addWorksheet(nombre, { views: [{ state: "frozen", ySplit: 1 }] });
    ws.columns = cols.map(function (c) { return { header: c[0], key: c[1], width: c[2], style: c[3] ? { numFmt: c[3] } : {} }; });
    filas.forEach(function (f) { ws.addRow(f); });
    encabezado(ws, 1);
    ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: cols.length } };
    return ws;
  }

  function titulo(ws, fila, texto, nota) {
    var c = ws.getCell(fila, 1);
    c.value = texto;
    c.font = { bold: true, size: 13, color: { argb: AZUL } };
    if (nota) {
      var n = ws.getCell(fila + 1, 1);
      n.value = nota;
      n.font = { italic: true, color: { argb: "FF5B6575" } };
      return fila + 2;
    }
    return fila + 1;
  }

  function cab(ws, fila, textos) {
    textos.forEach(function (t, i) { ws.getCell(fila, i + 1).value = t; });
    ws.getRow(fila).eachCell(function (c) {
      c.font = { bold: true };
      c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: CLARO } };
      c.alignment = { wrapText: true, vertical: "middle" };
    });
  }

  function col(n) { var s = ""; n++; while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }

  function construir(ExcelJS, d) {
    var wb = new ExcelJS.Workbook();
    wb.creator = "Generador de evidencias TSP";
    wb.created = new Date();
    wb.calcProperties.fullCalcOnLoad = true;
    var H = "0.00", F = "yyyy-mm-dd", P = "0.0";

    var wsRes = wb.addWorksheet("Resumen");

    // Semanas
    var wsSem = hojaDatos(wb, "Semanas", [["Hito", "hito", 22], ["Inicio", "ini", 14, F], ["Fin", "fin", 14, F]],
      d.hitos.map(function (h) { return { hito: h.titulo, ini: h.inicio, fin: h.fin }; }));
    var nSem = d.hitos.length;
    function semana(celda) {
      return { formula: "IF(" + celda + "=\"\",\"\",IFERROR(INDEX(Semanas!$A:$A,MATCH(" + celda + ",Semanas!$B:$B,1)),\"\"))" };
    }

    // Elementos
    var wsEl = hojaDatos(wb, "Elementos", [
      ["#", "iid", 6], ["Tipo", "tipo", 8], ["Título", "titulo", 44], ["Estado", "estado", 10],
      ["Etiquetas", "etiquetas", 30], ["Clase", "clase", 16], ["Requisito", "req", 34],
      ["Responsable", "resp", 24], ["Hito", "hito", 14], ["Estimado (h)", "est", 11, H],
      ["Dedicado (h)", "ded", 11, H], ["Creado", "creado", 12, F], ["Cerrado", "cerrado", 12, F],
      ["Semana de cierre", "semc", 14], ["Valor planeado (%)", "pv", 12, H], ["Valor ganado (%)", "ev", 12, H]
    ], d.items.map(function (it) {
      return { iid: it.iid, tipo: it.tipo, titulo: it.titulo, estado: it.estado, etiquetas: it.etiquetas.join(", "),
        clase: it.clase, req: it.padreIid ? "#" + it.padreIid + " " + it.padreTitulo : "", resp: it.responsable,
        hito: it.hito, est: it.estimado, ded: it.dedicado, creado: it.creado, cerrado: it.cerrado };
    }));
    var nEl = d.items.length;
    for (var r = 2; r <= nEl + 1; r++) {
      wsEl.getCell("N" + r).value = semana("M" + r);
      wsEl.getCell("O" + r).value = { formula: "IF(AND(F" + r + "=\"Tarea\",I" + r + "<>\"\"),J" + r + "/SUMIFS($J:$J,$F:$F,\"Tarea\",$I:$I,\"<>\")*100,0)" };
      wsEl.getCell("O" + r).numFmt = H;
      wsEl.getCell("P" + r).value = { formula: "IF(D" + r + "=\"Cerrado\",O" + r + ",0)" };
      wsEl.getCell("P" + r).numFmt = H;
    }

    // Tiempo
    var filasT = d.tiempo.map(function (t) {
      return { fecha: t.fecha, persona: t.persona, usuario: t.usuario, elemento: t.elemento, titulo: t.titulo,
        clase: t.clase, horas: t.horas, fase: t.fase || "Sin fase", texto: t.texto };
    });
    var wsT = hojaDatos(wb, "Tiempo", [
      ["Fecha", "fecha", 12, F], ["Semana", "sem", 14], ["Persona", "persona", 24], ["Usuario", "usuario", 16],
      ["Elemento", "elemento", 10], ["Título del elemento", "titulo", 40], ["Clase del elemento", "clase", 18],
      ["Horas", "horas", 8, H], ["Fase", "fase", 18], ["Texto registrado", "texto", 30]
    ], filasT);
    var nT = Math.max(filasT.length, d.modo === "csv" ? 300 : 0);
    for (r = 2; r <= nT + 1; r++) {
      wsT.getCell("B" + r).value = semana("A" + r);
      wsT.getCell("G" + r).value = { formula: "IF(LEFT(E" + r + ",1)=\"#\",IFERROR(INDEX(Elementos!$F:$F,MATCH(VALUE(MID(E" + r + ",2,10)),Elementos!$A:$A,0)),\"\"),IF(LEFT(E" + r + ",1)=\"!\",\"Solicitud de fusión\",\"\"))" };
    }
    if (d.modo === "csv") for (r = 2; r <= nT + 1; r++) { wsT.getCell("A" + r).numFmt = F; wsT.getCell("H" + r).numFmt = H; }

    // Fusiones
    var wsF = hojaDatos(wb, "Fusiones", [
      ["!", "iid", 6], ["Título", "titulo", 36], ["Autor", "autor", 22], ["Estado", "estado", 11],
      ["Tarea que cierra", "tarea", 12], ["Revisor", "revisor", 22], ["Aprobada por", "apr", 22],
      ["Fusionada", "fus", 12, F], ["Semana", "sem", 14], ["Líneas agregadas", "add", 10],
      ["Líneas eliminadas", "del", 10], ["Comentarios en líneas", "com", 11], ["Inspectores", "insp", 10],
      ["Horas de inspección", "hi", 11, H], ["Tasa de inspección (líneas por hora por inspector)", "tasa", 16, "0"]
    ], d.fusiones.map(function (f) {
      return { iid: f.iid, titulo: f.titulo, autor: f.autor, estado: f.estado, tarea: f.tarea, revisor: f.revisor,
        apr: f.aprobadaPor, fus: f.fusionada, add: f.agregadas, del: f.eliminadas, com: f.comentarios,
        insp: f.inspectores, hi: f.horasInspeccion };
    }));
    var nF = Math.max(d.fusiones.length, d.modo === "csv" ? 60 : 0);
    for (r = 2; r <= nF + 1; r++) {
      wsF.getCell("I" + r).value = semana("H" + r);
      wsF.getCell("O" + r).value = { formula: "IF(AND(N" + r + ">0,M" + r + ">0),J" + r + "/(N" + r + "/M" + r + "),\"\")" };
      wsF.getCell("O" + r).numFmt = "0";
      if (d.modo === "csv") { wsF.getCell("H" + r).numFmt = F; wsF.getCell("N" + r).numFmt = H; }
    }

    // Defectos
    var wsD = hojaDatos(wb, "Defectos", [
      ["#", "iid", 6], ["Título", "titulo", 40], ["Estado", "estado", 10], ["Tipo", "tipo", 18],
      ["Fase inyectada", "iny", 18], ["Fase removida", "rem", 18], ["Responsable", "resp", 22],
      ["Hito", "hito", 14], ["Horas de corrección", "h", 11, H], ["Registrado", "creado", 12, F]
    ], d.defectos.map(function (x) {
      return { iid: x.it.iid, titulo: x.it.titulo, estado: x.it.estado, tipo: x.tipo || "Sin tipo",
        iny: x.inyectado || "Sin fase", rem: x.removido || "Sin fase", resp: x.it.responsable, hito: x.it.hito,
        h: x.it.dedicado, creado: x.it.creado };
    }));

    // Riesgos
    var wsR = hojaDatos(wb, "Riesgos", [
      ["#", "iid", 6], ["Título", "titulo", 44], ["Impacto", "imp", 10], ["Probabilidad", "prob", 12],
      ["Prioridad", "pri", 12], ["Responsable", "resp", 22], ["Estado", "estado", 10]
    ], d.riesgos.map(function (x) {
      return { iid: x.it.iid, titulo: x.it.titulo, imp: x.impacto, prob: x.probabilidad, resp: x.it.responsable, estado: x.it.estado };
    }));
    for (r = 2; r <= d.riesgos.length + 1; r++) {
      wsR.getCell("E" + r).value = { formula: "IF(AND(C" + r + "=\"Alto\",D" + r + "=\"Alta\"),\"Primera\",IF(AND(C" + r + "=\"Alto\",D" + r + "=\"Media\"),\"Segunda\",IF(AND(C" + r + "=\"Medio\",D" + r + "=\"Alta\"),\"Tercera\",IF(AND(C" + r + "=\"Medio\",D" + r + "=\"Media\"),\"Cuarta\",\"Sin plan\"))))" };
    }

    // Tablas
    var wsTb = wb.addWorksheet("Tablas");
    wsTb.columns = [{ width: 30 }].concat(new Array(14).fill(0).map(function () { return { width: 15 }; }));
    var f = 1, i, ini, fin;
    var pos = {};

    f = titulo(wsTb, f, "Horas por fase", "Suma de la hoja Tiempo, por la fase escrita en cada registro.");
    cab(wsTb, f, ["Fase", "Horas", "% del total"]); f++;
    ini = f;
    FASES.concat(["Sin fase"]).forEach(function (fa) {
      wsTb.getCell(f, 1).value = fa;
      wsTb.getCell(f, 2).value = { formula: "SUMIFS(Tiempo!$H:$H,Tiempo!$I:$I,A" + f + ")" };
      f++;
    });
    fin = f - 1;
    wsTb.getCell(f, 1).value = "Total"; wsTb.getCell(f, 1).font = { bold: true };
    wsTb.getCell(f, 2).value = { formula: "SUM(B" + ini + ":B" + fin + ")" };
    for (i = ini; i <= fin; i++) wsTb.getCell(i, 3).value = { formula: "IF($B$" + f + "=0,0,B" + i + "/$B$" + f + "*100)" };
    for (i = ini; i <= f; i++) { wsTb.getCell(i, 2).numFmt = H; wsTb.getCell(i, 3).numFmt = P; }
    pos.horasTotal = "Tablas!B" + f;
    f += 2;

    f = titulo(wsTb, f, "Estimado frente a dedicado, por persona", "Estimado: suma de las estimaciones de las tareas asignadas a la persona. Dedicado: horas registradas por la persona.");
    cab(wsTb, f, ["Persona", "Estimado en sus tareas (h)", "Dedicado en tareas (h)", "Diferencia (h)", "Desviación (%)", "Dedicado total, incluidas reuniones y defectos (h)"]); f++;
    ini = f;
    d.personas.forEach(function (p) {
      wsTb.getCell(f, 1).value = p;
      wsTb.getCell(f, 2).value = { formula: "SUMIFS(Elementos!$J:$J,Elementos!$H:$H,A" + f + ",Elementos!$F:$F,\"Tarea\")" };
      wsTb.getCell(f, 3).value = { formula: "SUMIFS(Tiempo!$H:$H,Tiempo!$C:$C,A" + f + ",Tiempo!$G:$G,\"Tarea\")" };
      wsTb.getCell(f, 4).value = { formula: "C" + f + "-B" + f };
      wsTb.getCell(f, 5).value = { formula: "IF(B" + f + "=0,\"\",(C" + f + "-B" + f + ")/B" + f + "*100)" };
      wsTb.getCell(f, 6).value = { formula: "SUMIFS(Tiempo!$H:$H,Tiempo!$C:$C,A" + f + ")" };
      [2, 3, 4, 6].forEach(function (c) { wsTb.getCell(f, c).numFmt = H; });
      wsTb.getCell(f, 5).numFmt = P;
      f++;
    });
    f += 1;

    f = titulo(wsTb, f, "Estimado frente a dedicado, por semana", "Estimado: tareas del hito de la semana. Dedicado: horas registradas en las fechas de la semana.");
    cab(wsTb, f, ["Semana", "Estimado (h)", "Dedicado (h)", "Diferencia (h)"]); f++;
    d.hitos.forEach(function (h) {
      wsTb.getCell(f, 1).value = h.titulo;
      wsTb.getCell(f, 2).value = { formula: "SUMIFS(Elementos!$J:$J,Elementos!$I:$I,A" + f + ",Elementos!$F:$F,\"Tarea\")" };
      wsTb.getCell(f, 3).value = { formula: "SUMIFS(Tiempo!$H:$H,Tiempo!$B:$B,A" + f + ")" };
      wsTb.getCell(f, 4).value = { formula: "C" + f + "-B" + f };
      [2, 3, 4].forEach(function (c) { wsTb.getCell(f, c).numFmt = H; });
      f++;
    });
    f += 1;

    f = titulo(wsTb, f, "Horas dedicadas por persona y por semana", "Horas registradas por cada persona en las fechas de cada semana.");
    cab(wsTb, f, ["Persona"].concat(d.hitos.map(function (h) { return h.titulo; })).concat(["Total"])); var fh = f; f++;
    d.personas.forEach(function (p) {
      wsTb.getCell(f, 1).value = p;
      d.hitos.forEach(function (h, k) {
        wsTb.getCell(f, k + 2).value = { formula: "SUMIFS(Tiempo!$H:$H,Tiempo!$C:$C,$A" + f + ",Tiempo!$B:$B," + col(k + 1) + "$" + fh + ")" };
        wsTb.getCell(f, k + 2).numFmt = H;
      });
      wsTb.getCell(f, nSem + 2).value = { formula: "SUM(B" + f + ":" + col(nSem) + f + ")" };
      wsTb.getCell(f, nSem + 2).numFmt = H;
      f++;
    });
    f += 1;

    f = titulo(wsTb, f, "Valor planeado frente a valor ganado", "Valor planeado: tareas del hito de la semana. Valor ganado: tareas cerradas en la semana. Cada tarea vale su estimación sobre el total estimado del plan.");
    cab(wsTb, f, ["Semana", "Valor planeado (%)", "Valor ganado (%)", "Valor planeado acumulado (%)", "Valor ganado acumulado (%)", "Estado acumulado"]); f++;
    ini = f;
    d.hitos.forEach(function (h) {
      wsTb.getCell(f, 1).value = h.titulo;
      wsTb.getCell(f, 2).value = { formula: "SUMIFS(Elementos!$O:$O,Elementos!$I:$I,A" + f + ")" };
      wsTb.getCell(f, 3).value = { formula: "SUMIFS(Elementos!$P:$P,Elementos!$N:$N,A" + f + ")" };
      wsTb.getCell(f, 4).value = { formula: "SUM(B$" + ini + ":B" + f + ")" };
      wsTb.getCell(f, 5).value = { formula: "SUM(C$" + ini + ":C" + f + ")" };
      wsTb.getCell(f, 6).value = { formula: "IF(E" + f + ">=D" + f + ",\"Al día o adelantado\",\"Atrasado\")" };
      [2, 3, 4, 5].forEach(function (c) { wsTb.getCell(f, c).numFmt = P; });
      f++;
    });
    pos.evTotal = "SUM(Elementos!P:P)";
    f += 1;

    f = titulo(wsTb, f, "Defectos por tipo", "Conteo de la hoja Defectos.");
    cab(wsTb, f, ["Tipo", "Defectos"]); f++;
    ini = f;
    TIPOS.concat(["Sin tipo"]).forEach(function (t) {
      wsTb.getCell(f, 1).value = t;
      wsTb.getCell(f, 2).value = { formula: "COUNTIFS(Defectos!$D:$D,A" + f + ")" };
      f++;
    });
    wsTb.getCell(f, 1).value = "Total"; wsTb.getCell(f, 1).font = { bold: true };
    wsTb.getCell(f, 2).value = { formula: "SUM(B" + ini + ":B" + (f - 1) + ")" };
    f += 2;

    f = titulo(wsTb, f, "Defectos por fase inyectada (filas) y fase removida (columnas)", "Conteo de la hoja Defectos.");
    var fasesM = FASES.concat(["Sin fase"]);
    cab(wsTb, f, ["Inyectado \\ Removido"].concat(fasesM).concat(["Total"])); fh = f; f++;
    ini = f;
    fasesM.forEach(function (fa) {
      wsTb.getCell(f, 1).value = fa;
      fasesM.forEach(function (fb, k) {
        wsTb.getCell(f, k + 2).value = { formula: "COUNTIFS(Defectos!$E:$E,$A" + f + ",Defectos!$F:$F," + col(k + 1) + "$" + fh + ")" };
      });
      wsTb.getCell(f, fasesM.length + 2).value = { formula: "SUM(B" + f + ":" + col(fasesM.length) + f + ")" };
      f++;
    });
    wsTb.getCell(f, 1).value = "Total"; wsTb.getCell(f, 1).font = { bold: true };
    for (i = 2; i <= fasesM.length + 2; i++) wsTb.getCell(f, i).value = { formula: "SUM(" + col(i - 1) + ini + ":" + col(i - 1) + (f - 1) + ")" };
    f += 2;

    f = titulo(wsTb, f, "Indicadores de calidad", "Densidad: defectos por cada mil líneas agregadas. Tasa: líneas inspeccionadas por hora por inspector; TSP recomienda máximo 150.");
    cab(wsTb, f, ["Indicador", "Valor"]); f++;
    var ind = [
      ["Líneas agregadas en solicitudes de fusión", "SUM(Fusiones!$J:$J)", "0"],
      ["Defectos registrados", "COUNTA(Defectos!$A:$A)-1", "0"],
      ["Densidad de defectos (defectos por mil líneas)", "IF(B" + f + "=0,\"\",B" + (f + 1) + "/(B" + f + "/1000))", P],
      ["Tasa de inspección promedio (líneas por hora por inspector)", "IFERROR(AVERAGE(Fusiones!$O:$O),\"\")", "0"],
      ["Solicitudes de fusión con más de 150 líneas por hora", "COUNTIFS(Fusiones!$O:$O,\">150\")", "0"],
      ["Horas de corrección de defectos", "SUM(Defectos!$I:$I)", H]
    ];
    pos.ind = f;
    ind.forEach(function (x) {
      wsTb.getCell(f, 1).value = x[0];
      wsTb.getCell(f, 2).value = { formula: x[1] };
      wsTb.getCell(f, 2).numFmt = x[2];
      f++;
    });

    // Línea de tiempo
    var wsL = wb.addWorksheet("Línea de tiempo", { views: [{ state: "frozen", ySplit: 1 }] });
    wsL.columns = [
      { header: "Semana", width: 14 }, { header: "Inicio", width: 12, style: { numFmt: F } }, { header: "Fin", width: 12, style: { numFmt: F } },
      { header: "Tareas terminadas", width: 50 }, { header: "Requisitos completados", width: 40 },
      { header: "Horas del equipo", width: 11, style: { numFmt: H } }, { header: "Reuniones semanales", width: 11 },
      { header: "Relanzamientos", width: 11 }, { header: "Solicitudes fusionadas", width: 11 }, { header: "Defectos registrados", width: 11 },
      { header: "Riesgos registrados", width: 11 }
    ];
    encabezado(wsL, 1);
    function semDe(fecha) {
      if (!fecha) return "";
      var s = "";
      d.hitos.forEach(function (h) { if (h.inicio && fecha >= h.inicio) s = h.titulo; });
      return s;
    }
    d.hitos.forEach(function (h, k) {
      var rr = k + 2;
      var term = d.items.filter(function (it) { return it.clase === "Tarea" && it.estado === "Cerrado" && semDe(it.cerrado) === h.titulo; });
      var reqs = d.items.filter(function (it) {
        if (it.clase !== "Requisito" || it.estado !== "Cerrado") return false;
        return semDe(it.cerrado) === h.titulo;
      });
      wsL.getCell(rr, 1).value = h.titulo;
      wsL.getCell(rr, 2).value = { formula: "IFERROR(INDEX(Semanas!$B:$B,MATCH(A" + rr + ",Semanas!$A:$A,0)),\"\")" };
      wsL.getCell(rr, 3).value = { formula: "IFERROR(INDEX(Semanas!$C:$C,MATCH(A" + rr + ",Semanas!$A:$A,0)),\"\")" };
      wsL.getCell(rr, 4).value = term.map(function (it) { return it.titulo; }).join("; ");
      wsL.getCell(rr, 5).value = reqs.map(function (it) { return it.titulo; }).join("; ");
      wsL.getCell(rr, 6).value = { formula: "SUMIFS(Tiempo!$H:$H,Tiempo!$B:$B,A" + rr + ")" };
      wsL.getCell(rr, 7).value = { formula: "COUNTIFS(Elementos!$F:$F,\"Reunión semanal\",Elementos!$I:$I,A" + rr + ")" };
      wsL.getCell(rr, 8).value = { formula: "COUNTIFS(Elementos!$F:$F,\"Relanzamiento\",Elementos!$I:$I,A" + rr + ")" };
      wsL.getCell(rr, 9).value = { formula: "COUNTIFS(Fusiones!$I:$I,A" + rr + ",Fusiones!$D:$D,\"Fusionada\")" };
      wsL.getCell(rr, 10).value = { formula: "COUNTIFS(Elementos!$F:$F,\"Defecto\",Elementos!$I:$I,A" + rr + ")" };
      wsL.getCell(rr, 11).value = { formula: "COUNTIFS(Elementos!$F:$F,\"Riesgo\",Elementos!$I:$I,A" + rr + ")" };
      [4, 5].forEach(function (c) { wsL.getCell(rr, c).alignment = { wrapText: true, vertical: "top" }; });
    });

    // Análisis
    var wsA = wb.addWorksheet("Análisis");
    wsA.columns = [{ header: "Tabla", width: 40 }, { header: "Lo que muestran los datos", width: 100 }];
    encabezado(wsA, 1);
    ["Horas por fase", "Estimado frente a dedicado, por persona", "Estimado frente a dedicado, por semana",
      "Horas dedicadas por persona y por semana", "Valor planeado frente a valor ganado", "Defectos por tipo",
      "Defectos por fase inyectada y fase removida", "Indicadores de calidad", "Riesgos", "Línea de tiempo"].forEach(function (t, k) {
      wsA.getCell(k + 2, 1).value = t;
      wsA.getCell(k + 2, 1).alignment = { vertical: "top" };
      wsA.getCell(k + 2, 2).alignment = { wrapText: true, vertical: "top" };
      wsA.getRow(k + 2).height = 90;
    });

    // Resumen
    wsRes.columns = [{ width: 46 }, { width: 60 }];
    wsRes.getCell("A1").value = "Datos del proyecto";
    wsRes.getCell("A1").font = { bold: true, size: 15, color: { argb: AZUL } };
    var filasR = [
      ["Proyecto", d.proyecto.nombre || ""],
      ["Dirección", d.proyecto.url || ""],
      ["Origen de los datos", d.modo === "gitlab" ? "GitLab, consulta directa" : "CSV de elementos de trabajo"],
      ["Fecha de generación", new Date(), F],
      ["Integrantes con registros", d.personas.length],
      ["Horas registradas por el equipo", { formula: "SUM(Tiempo!$H:$H)" }, H],
      ["Requisitos", { formula: "COUNTIF(Elementos!$F:$F,\"Requisito\")" }],
      ["Tareas del plan", { formula: "COUNTIFS(Elementos!$F:$F,\"Tarea\",Elementos!$I:$I,\"<>\")" }],
      ["Tareas terminadas", { formula: "COUNTIFS(Elementos!$F:$F,\"Tarea\",Elementos!$I:$I,\"<>\",Elementos!$D:$D,\"Cerrado\")" }],
      ["Horas estimadas en el plan", { formula: "SUMIFS(Elementos!$J:$J,Elementos!$F:$F,\"Tarea\",Elementos!$I:$I,\"<>\")" }, H],
      ["Valor ganado acumulado (%)", { formula: "SUM(Elementos!$P:$P)" }, P],
      ["Reuniones semanales", { formula: "COUNTIF(Elementos!$F:$F,\"Reunión semanal\")" }],
      ["Solicitudes de fusión fusionadas", { formula: "COUNTIF(Fusiones!$D:$D,\"Fusionada\")" }],
      ["Líneas de código agregadas", { formula: "SUM(Fusiones!$J:$J)" }],
      ["Defectos registrados", { formula: "COUNTA(Defectos!$A:$A)-1" }],
      ["Densidad de defectos (defectos por mil líneas)", { formula: "Tablas!B" + (pos.ind + 2) }, P],
      ["Riesgos registrados", { formula: "COUNTA(Riesgos!$A:$A)-1" }]
    ];
    filasR.forEach(function (x, k) {
      var rr = k + 3;
      wsRes.getCell(rr, 1).value = x[0];
      wsRes.getCell(rr, 1).font = { bold: true };
      wsRes.getCell(rr, 2).value = x[1];
      wsRes.getCell(rr, 2).alignment = { horizontal: "left" };
      if (x[2]) wsRes.getCell(rr, 2).numFmt = x[2];
      wsRes.getCell(rr, 1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: GRIS } };
    });
    var rr0 = filasR.length + 4;
    if (d.avisos.length) {
      wsRes.getCell(rr0, 1).value = "Avisos sobre los datos";
      wsRes.getCell(rr0, 1).font = { bold: true, color: { argb: "FF8A3B12" } };
      d.avisos.forEach(function (a, k) {
        wsRes.mergeCells(rr0 + 1 + k, 1, rr0 + 1 + k, 2);
        var c = wsRes.getCell(rr0 + 1 + k, 1);
        c.value = a;
        c.alignment = { wrapText: true, vertical: "top" };
        wsRes.getRow(rr0 + 1 + k).height = 32;
      });
    }
    return wb;
  }

  var api = { FASES: FASES, TIPOS: TIPOS, fase: fase, horas: horas, cierres: cierres, rutaProyecto: rutaProyecto,
    parseCSV: parseCSV, leerCSV: leerCSV, leerGitLab: leerGitLab, completar: completar, construir: construir,
    resumen: resumen, unir: unir, fechaLocal: fechaLocal, tipoDefecto: tipoDefecto, faseEtiqueta: faseEtiqueta };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.TSPGen = api;
})(this);
