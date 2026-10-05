class AuraHero extends HTMLElement {
    connectedCallback() {
        this.innerHTML = /*html*/`
        <header class="splash-hero" id="inicio">
            <div class="splash-content">
                <img src="assets/images/logo_dark.png" alt="AURA Soluciones e Innovación" class="splash-logo">
            </div>
            
            <a href="#intro" class="scroll-indicator">
                <span style="display: block; font-size: 0.95rem; color: var(--text-muted); margin-bottom: 8px; font-weight: 600;">Conoce más</span>
                <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="var(--accent-primary)" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round" class="bounce"><path d="M6 9l6 6 6-6"/></svg>
            </a>
            
            <img src="assets/images/mountain_dark.png" alt="" class="bg-mountain-massive">
        </header>

        <section id="intro" class="intro-section wrap" style="padding-top: 100px; padding-bottom: 100px;">
            <div class="eyebrow" style="text-align: center;">Aura Soluciones e Innovación S.A.S.</div>
            <h1 style="text-align: center; color: var(--text-main); font-size: clamp(36px, 5vw, 64px); line-height: 1.1; margin: 1.5rem auto; max-width: 900px;">Tecnología que convierte ideas en soluciones.</h1>
            <p class="lead" style="text-align: center; margin: 0 auto 3rem; font-size: 1.25rem;">Desarrollamos software a la medida y productos SaaS para organizaciones que buscan evolucionar, mejorar sus procesos y crear mejores experiencias digitales.</p>
            <div class="buttons" style="justify-content: center;">
                <a class="btn btn-primary" href="#servicios">Conoce Aura</a>
                <a class="btn btn-secondary" style="color: var(--text-main); border-color: var(--text-main);" href="#skooling">Conoce Skooling</a>
            </div>
        </section>
        `;
    }
}
customElements.define('aura-hero', AuraHero);
