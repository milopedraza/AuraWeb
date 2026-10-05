class AuraTecnologia extends HTMLElement {
    connectedCallback() {
        this.innerHTML = /*html*/`
        <section id="tecnologia">
            <div class="wrap">
                <div class="eyebrow">Tecnología</div>
                <h2>Construimos con tecnología pensada para durar.</h2>
                <p class="lead">Ingeniería, arquitectura y experiencia de usuario para crear soluciones útiles, mantenibles y preparadas para evolucionar.</p>
                <div class="grid">
                    <div class="card">
                        <h3>Aplicaciones web</h3>
                        <p>Experiencias digitales para usuarios y equipos.</p>
                    </div>
                    <div class="card">
                        <h3>Aplicaciones móviles</h3>
                        <p>Productos pensados para acompañar al usuario donde los necesita.</p>
                    </div>
                    <div class="card">
                        <h3>Servicios cloud</h3>
                        <p>Infraestructura y servicios orientados a soluciones modernas.</p>
                    </div>
                </div>
            </div>
        </section>
        `;
    }
}
customElements.define('aura-tecnologia', AuraTecnologia);
