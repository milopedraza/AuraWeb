class AuraServicios extends HTMLElement {
    connectedCallback() {
        this.innerHTML = /*html*/`
        <section id="servicios">
            <div class="wrap">
                <div class="eyebrow">Qué hacemos</div>
                <h2>Soluciones digitales construidas para evolucionar.</h2>
                <p class="lead">Trabajamos en diferentes etapas de la transformación digital, desde la creación de nuevos productos hasta la evolución de sistemas existentes.</p>
                
                <div class="grid">
                    <div class="card">
                        <h3>Software a la medida</h3>
                        <p>Diseñamos y desarrollamos soluciones digitales adaptadas a las necesidades específicas de cada organización.</p>
                    </div>
                    <div class="card">
                        <h3>Software como servicio (SaaS)</h3>
                        <p>Creamos productos digitales que pueden evolucionar como servicios tecnológicos, orientados a usuarios y organizaciones.</p>
                    </div>
                    <div class="card">
                        <h3>Cloud y arquitectura</h3>
                        <p>Diseñamos soluciones modernas y escalables apoyadas en servicios cloud, buscando equilibrio entre rendimiento, mantenibilidad y seguridad.</p>
                    </div>
                    <div class="card">
                        <h3>Modernización</h3>
                        <p>Ayudamos a evolucionar sistemas existentes, integrando nuevas capacidades sin perder de vista las necesidades del negocio.</p>
                    </div>
                    <div class="card">
                        <h3>Automatización e IA</h3>
                        <p>Exploramos y desarrollamos soluciones de automatización e inteligencia artificial cuando aportan valor real al proceso o producto.</p>
                    </div>
                    <div class="card">
                        <h3>Consultoría tecnológica</h3>
                        <p>Acompañamos decisiones relacionadas con arquitectura, desarrollo y evolución de productos digitales.</p>
                    </div>
                </div>
            </div>
        </section>
        `;
    }
}
customElements.define('aura-servicios', AuraServicios);
