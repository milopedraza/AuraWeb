class AuraSocialProof extends HTMLElement {
    connectedCallback() {
        this.innerHTML = /*html*/`
        <section id="testimonios" class="pillars-section">
            <div class="section-header">
                <h2 style="color: white; margin-bottom: 3rem; text-align: center;">Colegios que ya dieron el paso</h2>
            </div>
            
            <div class="pillars-grid">
                <div class="pillar-box">
                    <h3>Colegio Piloto A</h3>
                    <p>"Migrar a Skooling fue el paso que necesitábamos. La confianza de los padres ha aumentado exponencialmente y el equipo administrativo finalmente respira."</p>
                    <p style="color: var(--accent-primary); margin-top: 1rem;">— Rector</p>
                </div>
                <div class="pillar-box">
                    <h3>Colegio Piloto B</h3>
                    <p>"El respaldo técnico de Aura es evidente. No se trata solo del software, sino de saber que hay ingenieros asegurando que todo funcione."</p>
                    <p style="color: var(--accent-primary); margin-top: 1rem;">— Director de TI</p>
                </div>
                <div class="pillar-box">
                    <h3>Colegio Piloto C</h3>
                    <p>"Pensamos que adoptar una nueva tecnología iba a generar mucha resistencia, pero Skooling fue tan simple de implementar que nadie quiso volver atrás."</p>
                    <p style="color: var(--accent-primary); margin-top: 1rem;">— Coordinadora Académica</p>
                </div>
            </div>
        </section>
        `;
    }
}
customElements.define('aura-social-proof', AuraSocialProof);
