class AuraHero extends HTMLElement {
    connectedCallback() {
        this.innerHTML = /*html*/`
        <header class="executive-hero" id="inicio">
            <div class="hero-glow" aria-hidden="true"></div>
            <img src="assets/images/mountain_light.png" alt="" class="hero-watermark" aria-hidden="true">

            <div class="hero-inner">
                <div class="eyebrow hero-eyebrow">Aura Soluciones e Innovación S.A.S.</div>
                <h1 class="hero-title">Tecnología que convierte ideas en <span>soluciones</span>.</h1>
                <p class="hero-lead">
                    Desarrollamos software a la medida y productos SaaS para organizaciones que buscan evolucionar, mejorar sus procesos y crear mejores experiencias digitales.
                </p>
                <div class="buttons hero-buttons">
                    <a class="btn btn-primary" href="#servicios">Conoce Aura</a>
                    <a class="btn btn-secondary" href="#skooling">Conoce Skooling</a>
                </div>

                <div class="hero-pillars">
                    <div class="hero-pillar">
                        <strong>Software a la Medida</strong>
                        <span>Plataformas web y móviles diseñadas para tu operación.</span>
                    </div>
                    <div class="hero-pillar">
                        <strong>Ecosistemas SaaS</strong>
                        <span>Productos propios escalables como Skooling.</span>
                    </div>
                    <div class="hero-pillar">
                        <strong>Infraestructura Cloud</strong>
                        <span>Arquitectura segura, moderna y de alta disponibilidad.</span>
                    </div>
                </div>
            </div>
        </header>
        `;
    }
}
customElements.define('aura-hero', AuraHero);
