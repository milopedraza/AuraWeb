class AuraSkooling extends HTMLElement {
    connectedCallback() {
        this.innerHTML = /*html*/`
        <section id="skooling" class="skool section-alt">
            <div class="wrap" style="padding: 100px 5%;">
                <div class="product">
                    <div>
                        <div class="eyebrow">Producto SaaS de Aura</div>
                        <h2>La vida escolar de tu hijo, en un solo lugar.</h2>
                        <p class="lead">Skooling es una plataforma digital que conecta a las familias con la información y las actividades que forman parte de la experiencia escolar de sus hijos.</p>
                        <div class="buttons">
                            <a class="btn btn-primary" href="#contacto">Conocer Skooling</a>
                        </div>
                    </div>
                    <div class="quote">
                        “Una sola aplicación para todo lo que importa de la vida escolar.”
                    </div>
                </div>
                <div class="product-features" style="margin-top: 5rem;">
                    <div class="feature-row">
                        <div class="feature-text">
                            <h3>Experiencia académica</h3>
                            <p class="lead" style="margin-bottom:0;">Horarios, clases, tareas, actividades, eventos, calificaciones y otros elementos de seguimiento que la institución habilite.</p>
                        </div>
                        <div class="feature-image">
                            <div class="mockup-web">
                                <div class="mockup-header">
                                    <i></i><i></i><i></i>
                                </div>
                                <img src="assets/images/skooling_real_academica.png" alt="Skooling Académica">
                            </div>
                        </div>
                    </div>

                    <div class="feature-row reverse">
                        <div class="feature-text">
                            <h3>Comunicación</h3>
                            <p class="lead" style="margin-bottom:0;">Conexión entre familias, docentes e institución dentro de una misma experiencia digital.</p>
                        </div>
                        <div class="feature-image">
                            <div class="mockup-mobile">
                                <div class="notch"></div>
                                <img src="assets/images/skooling_real_comunicacion.png" alt="Skooling Comunicación">
                            </div>
                        </div>
                    </div>

                    <div class="feature-row">
                        <div class="feature-text">
                            <h3>Movilidad</h3>
                            <p class="lead" style="margin-bottom:0;">Cuando existe el servicio de transporte, información relacionada con la ruta escolar y su ubicación.</p>
                        </div>
                        <div class="feature-image">
                            <div class="mockup-web">
                                <div class="mockup-header">
                                    <i></i><i></i><i></i>
                                </div>
                                <img src="assets/images/skooling_real_movilidad.png" alt="Skooling Movilidad">
                            </div>
                        </div>
                    </div>
                </div>
                <div class="steps">
                    <div class="step">
                        <strong>Antes de la jornada</strong>
                        ¿Qué tiene hoy mi hijo?
                    </div>
                    <div class="step">
                        <strong>Durante el día</strong>
                        ¿Qué está pasando?
                    </div>
                    <div class="step">
                        <strong>Seguimiento</strong>
                        ¿Cómo va?
                    </div>
                    <div class="step">
                        <strong>Comunicación</strong>
                        ¿Necesito hablar con el colegio?
                    </div>
                    <div class="step">
                        <strong>Regreso</strong>
                        ¿Ya viene de camino?
                    </div>
                </div>
            </div>
        </section>
        `;
    }
}
customElements.define('aura-skooling', AuraSkooling);
