class AuraNosotros extends HTMLElement {
    connectedCallback() {
        this.innerHTML = /*html*/`
        <section id="nosotros" class="section-alt">
            <div class="wrap" style="padding: 100px 5%;">
                <div class="eyebrow">Aura</div>
                <h2>Tecnología con visión de largo plazo.</h2>
                <p class="lead">Aura Soluciones e Innovación S.A.S. busca convertir la tecnología en un motor de crecimiento y confianza para organizaciones que quieren transformarse digitalmente.</p>
                <div class="grid">
                    <div class="card">
                        <h3>Confianza</h3>
                        <p>Relaciones basadas en claridad, responsabilidad y respaldo técnico.</p>
                    </div>
                    <div class="card">
                        <h3>Resultados profesionales</h3>
                        <p>Soluciones bien diseñadas y alineadas con los objetivos del proyecto.</p>
                    </div>
                    <div class="card">
                        <h3>Evolución</h3>
                        <p>La tecnología debe poder crecer junto con las necesidades del negocio.</p>
                    </div>
                </div>
                
                <div style="margin-top: 4rem; padding: 3rem; background: var(--bg-color); border-radius: 16px; border: 1px solid var(--line); box-shadow: 0 4px 20px rgba(0, 0, 0, 0.02);">
                    <h3 style="margin-bottom: 1rem; color: var(--primary);">Experiencia internacional que respalda nuestra tecnología</h3>
                    <p style="color: var(--text-muted); line-height: 1.8; margin: 0;">
                        Aura nace y se fundamenta en la experiencia práctica real en ingeniería de software y construcción de soluciones tecnológicas. Combinamos visión estratégica de producto con más de 15 años de trayectoria corporativa demostrable construyendo soluciones para empresas tecnológicas en <strong>Norteamérica, Europa y Suramérica</strong>. Esta exposición a mercados globales garantiza la adopción de los más altos estándares internacionales en todas las fases del ciclo de vida del software: desde la especificación rigurosa y diseño de arquitectura, hasta el desarrollo, automatización de pruebas y evolución de productos digitales complejos.
                    </p>
                </div>
            </div>
        </section>
        `;
    }
}
customElements.define('aura-nosotros', AuraNosotros);
