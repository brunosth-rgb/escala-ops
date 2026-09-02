import React from "react";
import { createRoot } from "react-dom/client";
import App from "./App.jsx";

/*
 * O app foi escrito para o window.storage do ambiente de artifacts.
 * Aqui recriamos a mesma interface antes de montar a tela, então o
 * código do quadro não precisou de nenhuma alteração.
 *
 *   compartilhado = true  -> vai para o servidor, todos veem
 *   compartilhado = false -> fica só neste navegador
 */

const url = (rota, params) =>
  `/api/${rota}?${new URLSearchParams(params).toString()}`;

const local = {
  get(chave) {
    const v = window.localStorage.getItem(`local:${chave}`);
    if (v === null) throw new Error("não encontrado");
    return { key: chave, value: v, shared: false };
  },
  set(chave, valor) {
    window.localStorage.setItem(`local:${chave}`, valor);
    return { key: chave, value: valor, shared: false };
  },
  delete(chave) {
    window.localStorage.removeItem(`local:${chave}`);
    return { key: chave, deleted: true, shared: false };
  },
  list(prefixo = "") {
    const chaves = [];
    for (let i = 0; i < window.localStorage.length; i++) {
      const k = window.localStorage.key(i);
      if (k?.startsWith(`local:${prefixo}`)) chaves.push(k.slice(6));
    }
    return { keys: chaves, prefix: prefixo, shared: false };
  },
};

window.storage = {
  async get(chave, compartilhado = false) {
    if (!compartilhado) return local.get(chave);
    const r = await fetch(url("valor", { chave }));
    if (r.status === 404) throw new Error("não encontrado");
    if (!r.ok) throw new Error(`servidor respondeu ${r.status}`);
    const { valor } = await r.json();
    return { key: chave, value: valor, shared: true };
  },

  async set(chave, valor, compartilhado = false) {
    if (!compartilhado) return local.set(chave, valor);
    const r = await fetch(url("valor", { chave }), {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ valor }),
    });
    if (!r.ok) throw new Error(`servidor respondeu ${r.status}`);
    return { key: chave, value: valor, shared: true };
  },

  async delete(chave, compartilhado = false) {
    if (!compartilhado) return local.delete(chave);
    const r = await fetch(url("valor", { chave }), { method: "DELETE" });
    if (!r.ok) throw new Error(`servidor respondeu ${r.status}`);
    return { key: chave, deleted: true, shared: true };
  },

  async list(prefixo = "", compartilhado = false) {
    if (!compartilhado) return local.list(prefixo);
    const r = await fetch(url("lista", { prefixo }));
    if (!r.ok) throw new Error(`servidor respondeu ${r.status}`);
    const { chaves } = await r.json();
    return { keys: chaves, prefix: prefixo, shared: true };
  },
};

createRoot(document.getElementById("raiz")).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);
