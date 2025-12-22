(async () => {

  try {

    // 1️⃣ Descobrir o site do SharePoint
    const ref = document.referrer;

    if (!ref) {
      throw "Iframe não está vindo do SharePoint";
    }

    const siteUrl = ref.split("/SitePages")[0];

    console.log("Site:", siteUrl);

    // 2️⃣ Pegar usuário logado
    const userResp = await fetch(
      `${siteUrl}/_api/web/currentuser`,
      { credentials: "include" }
    );

    if (!userResp.ok) {
      throw "Erro ao obter usuário";
    }

    const user = await userResp.json();

    console.log("Usuário:", user);

    document.getElementById("status").innerText =
      `Usuário: ${user.Title} (${user.Email})`;

    // 3️⃣ Buscar itens da lista
    const listName = "Solicitações";

    const listResp = await fetch(
      `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items?$filter=Author/Email eq '${user.Email}'&$expand=Author`,
      {
        headers: {
          "Accept": "application/json;odata=nometadata"
        },
        credentials: "include"
      }
    );

    if (!listResp.ok) {
      throw "Erro ao buscar lista";
    }

    const data = await listResp.json();

    renderizar(data.value);

  } catch (e) {
    console.error(e);
    document.getElementById("status").innerText =
      "Erro: veja o console (F12)";
  }

})();

function renderizar(itens) {

  if (!itens.length) {
    document.getElementById("resultado").innerText =
      "Nenhuma solicitação encontrada.";
    return;
  }

  let html = "<ul>";

  itens.forEach(i => {
    html += `<li>${i.Title}</li>`;
  });

  html += "</ul>";

  document.getElementById("resultado").innerHTML = html;
}
