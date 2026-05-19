const projects = [
  {
    title: "Hub estatico no GitHub Pages",
    description: "Estrutura base para publicar materiais, documentos e prototipos do time.",
    status: "ativo",
  },
  {
    title: "Base de respostas tecnicas",
    description: "Organizacao inicial para documentos de respostas a questionamentos.",
    status: "ativo",
  },
  {
    title: "Galeria de midias",
    description: "Area reservada para futuras imagens e videos de apoio.",
    status: "planejado",
  },
];

const projectGrid = document.querySelector("#project-grid");
const materialCount = document.querySelector("#material-count");
const menuToggle = document.querySelector(".menu-toggle");
const nav = document.querySelector("#site-nav");

function createProjectCard(project) {
  const article = document.createElement("article");
  article.className = "project-card";
  article.innerHTML = `
    <span class="status">${project.status}</span>
    <h3>${project.title}</h3>
    <p>${project.description}</p>
  `;

  return article;
}

projectGrid.replaceChildren(...projects.map(createProjectCard));
materialCount.textContent = "1";

menuToggle.addEventListener("click", () => {
  const isOpen = nav.classList.toggle("open");
  menuToggle.setAttribute("aria-expanded", String(isOpen));
});

nav.addEventListener("click", (event) => {
  if (event.target.matches("a")) {
    nav.classList.remove("open");
    menuToggle.setAttribute("aria-expanded", "false");
  }
});
