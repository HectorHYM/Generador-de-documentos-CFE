//* Se carga primeramente todo el DOM (Document Object Model)
document.addEventListener("DOMContentLoaded", () => {

    //* Se carga el historial de reportes al iniciar la página
    load_historial();

    //* Se obtiene la referencia de los botones en la interfaz
    const generatorBtn = document.getElementById('generatorBtn'); //* Boton para generar los documentos
    const cleanBtn = document.getElementById('cleanHistorialBtn'); //* Boton para limpiar el historial

    //* Se agrega un evento click al botón generador de documentos
    generatorBtn.addEventListener('click', function() {
        //* Se obtiene el mes seleccionado por el usuario a través de un elemento select
        const month = document.getElementById("monthSelect").value;

        //* Se muestra un modal de carga utilizando SweetAlert (Swal) para informar al usuario que el proceso esta en curso
        Swal.fire({
            title: 'Generando reportes...',
            html: `<p class="atkinson-hyperlegible-next">Por favor, espere mientras se generan los reportes.</p>`,
            allowOutsideClick: false, //? Evita que se cierre el modal haciendo click fuera de él
            didOpen: () => {
                Swal.showLoading() //? Se muestra una animación de carga
            },
            customClass: {
                title: "load-title atkinson-hyperlegible-next"
            }
        });

        //* Se realiza una petición HTTP de tipo POST a la ruta '/generate' del servidor Flask
        fetch("http://127.0.0.1:5000/generate", { 
            method: "POST", //? Método HTTP POST para enviar datos al servidor
            headers: {
                "Content-Type": "application/json" //? Especifica que se enviará JSON en la solicitud
            },
            body: JSON.stringify({ month: month }) //? Se envía el mes seleccionado en formato JSON
        })
        .then(response => response.json()) //? Se convierte la respuesta del servidor a un objeto JSON
        .then(data => {
            //? Se cierra el modal de carga (implícito al mostrar otro modal)
            //* Se muestra el modal de error en caso de que la respuesta contenga un mensaje de error
            if(data.error){
                Swal.fire({
                    title: "Error",
                    html: `<p class="atkinson-hyperlegible-next">${data.details || data.error}</p>`,
                    icon: "error",
                    customClass: {
                        title: "atkinson-hyperlegible-next"
                    }
                });
            } else{
                //* Si la generación fue exitosa, se extrae el resumen de la respuesta
                let summary = data.summary;
                //* Se muestra el modal con el resumen y el botón para la descarga del ZIP
                Swal.fire({
                    title: "Detalles de reportes",
                    html: `<p class="atkinson-hyperlegible-next"><strong>Mes procesado:</strong> ${summary.month}</p>
                           <p class="atkinson-hyperlegible-next"><strong>Total de documentos generados:</strong> ${summary.total_docs}</p>
                           <p class="atkinson-hyperlegible-next"><strong>Cursos sin participantes:</strong> ${summary.courses_without_participants}</p>`,
                    icon: "success",
                    confirmButtonText: "Descargar reportes A20",
                    customClass: {
                        popup: "details-alert",
                        title: "details-title atkinson-hyperlegible-next",
                        confirmButton: "details-btn atkinson-hyperlegible-next"
                    }
                }).then((result) => {
                    //? Si el usuario confirma (click en el botón de la generación de los reportes), se redirige a la URL de descarga
                    if(result.isConfirmed){
                        //? Se redirige a la URL para forzar la descarga del ZIP
                        window.location.href = summary.download_url;
                        //* Se actualiza el historial en tiempo real después de la confirmación exitosa
                        load_historial()
                    }
                });
            }
        })
        .catch(error => {
            //? En caso de error en la petición, se cierra el modal de carga y se muestra un mensaje de error
            Swal.close();
            Swal.fire({
                title: "Error",
                text: error.details || "Hubo un error al generar los reportes.",
                icon: "Error"
            });

            //* Se registra el error en la consola para depuración 
            console.error("Error: ", error);
        });
    });

    //* Se agrega un evento click al botón para limpiar el historial
    cleanBtn.addEventListener("click", function() {
        fetch("http://127.0.0.1:5000/clean_historial", { method: "POST"})
        .then(response => response.json())
        .then(data => {
            if(data.error){
                //* Se muestra un modal de error en caso de fallo al limpiar el historial
                Swal.fire({ 
                    title: "Error",
                    text: data.details || data.error,
                    icon: "error",
                    customClass: {
                        title: "atkinson-hyperlegible-next",
                        text: "atkinson-hyperlegible-next",
                        confirmButton: "details-btn atkinson-hyperlegible-next"
                    }});
            }else{
                //* Se muestra un modal de éxito al limpiar el historial
                Swal.fire({ 
                    title: "Exito", 
                    html: `<p class="atkinson-hyperlegible-next">${data.message}</p`,
                    icon: "success",
                    customClass: {
                        title: "atkinson-hyperlegible-next",
                        confirmButton: "details-btn atkinson-hyperlegible-next"
                    }});
                load_historial(); //* Se actualiza la lista del historial para mostrarla vacía
            }
        }).catch(error => {
            console.error("Error al limpiar el historial: ", error);
            Swal.fire({ title: "Error", text: "No fue posible limpiar el historial.", icon: "error"});
        })
    });
});

//* Función: load_historial
//* Objetivo: Cargar y mostrar en tiempo real la lista de documentos del historial
const load_historial = () => {
    fetch("http://127.0.0.1:5000/historial")
        .then(response => response.json())
        .then(data => {
            console.log("Historial cargado: ", data); //* Se verifica que se reciba la respuesta correctamente

            let list = document.getElementById("historial-list"); //* Se obtiene el contenedor donde se mostrará el historial
            if(!list){
                console.error("No se escontro el contenedor del historial en el HTML.");
                return;
            }

            list.innerHTML = ""; //* Se limpia la lista actual
            if(data.historial && data.historial.length > 0){
                //* Si existen archivos en el historial, se crean elementos de lista para cada uno
                data.historial.forEach((file) => {
                    let li = document.createElement("li");
                    li.className = "historial-file atkinson-hyperlegible-next";
                    li.textContent = file;
                    list.appendChild(li);
                });
            }else{
                //* En caso de no haber archivos se muestra un mensaje indicándolo
                list.innerHTML = `<li class="atkinson-hyperlegible-next">No hay reportes en el historial</li>`;
            }
        }).catch(error => console.error("Error al cargar el historial: ", error));
};