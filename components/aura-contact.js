class AuraContact extends HTMLElement {
    connectedCallback() {
        this.innerHTML = /*html*/`
        <section id="contacto">
            <div class="wrap">
                <div class="eyebrow">Contacto</div>
                <h2>Hablemos de su próximo proyecto.</h2>
                <p class="lead">Si busca desarrollar un producto digital, modernizar una solución existente o explorar una oportunidad tecnológica, conversemos.</p>
                
                <form id="contactForm" action="https://formsubmit.co/camilo@aurasi.co" method="POST" class="contact-form">
                    <input type="hidden" name="_captcha" value="false">
                    <input type="hidden" name="_subject" value="Nuevo prospecto corporativo desde AuraWeb">
                    
                    <div class="form-group">
                        <label for="name">Nombre Completo</label>
                        <input type="text" id="name" name="Nombre" required placeholder="Su nombre">
                    </div>
                    <div class="form-group">
                        <label for="company">Empresa / Institución</label>
                        <input type="text" id="company" name="Organizacion" required placeholder="Nombre de su organización">
                    </div>
                    <div class="form-group">
                        <label for="email">Correo Corporativo</label>
                        <input type="email" id="email" name="Email" required placeholder="correo@organizacion.com">
                    </div>
                    <div class="form-group">
                        <label for="message">Mensaje (Opcional)</label>
                        <textarea id="message" name="Mensaje" rows="4" placeholder="¿En qué podemos ayudarle?"></textarea>
                    </div>
                    
                    <button type="submit" class="btn btn-primary" style="width: 100%;">Contactar a Aura</button>
                </form>

                <div class="note">
                    <strong>Información Corporativa Oficial:</strong> Todos los enlaces, contactos comerciales y detalles adicionales están en proceso de configuración y verificación oficial por parte de Aura.
                </div>
            </div>
        </section>
        `;
    }
}
customElements.define('aura-contact', AuraContact);
