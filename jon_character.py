
import json
import streamlit as st
import streamlit.components.v1 as components


CHARACTER_VERSION = "10.0"


def render_character(
    height=620,
    name="JON",
    mode="idle",
    speaking=False,
    style="Human + Android",
    model_url=None,
):
    """JON AI 10.0 — Independent Character Engine."""

    config = {
        "name": name,
        "mode": "speaking" if speaking else mode,
        "style": style,
        "modelUrl": model_url or "",
    }

    html = r"""
<!DOCTYPE html>
<html>
<head>
<meta name="viewport" content="width=device-width, initial-scale=1">
<style>
* { box-sizing:border-box; }
body {
    margin:0;
    overflow:hidden;
    background:transparent;
    font-family:Arial,sans-serif;
}
#stage {
    width:100%;
    height:HEIGHTpx;
    position:relative;
    overflow:hidden;
    border-radius:20px;
    background:
      radial-gradient(ellipse at 50% 55%,
      #123c65 0%, #071629 48%, #020611 100%);
}
canvas { display:block; }
#info {
    position:absolute;
    top:16px;
    left:0;
    right:0;
    text-align:center;
    pointer-events:none;
    color:#65dfff;
    letter-spacing:4px;
    font-size:12px;
    text-shadow:0 0 15px #00aaff;
}
#status {
    position:absolute;
    bottom:17px;
    left:0;
    right:0;
    text-align:center;
    color:#e9c76e;
    font-size:12px;
    letter-spacing:3px;
}
#fallback {
    display:none;
    position:absolute;
    inset:0;
    align-items:center;
    justify-content:center;
    color:#8bdcff;
    text-align:center;
}
</style>
</head>
<body>
<div id="stage">
  <div id="info">JON AI • LIVE INTELLIGENCE</div>
  <div id="status">INITIALIZING JON...</div>
  <div id="fallback">JON AI<br>3D ENGINE LOADING...</div>
</div>

<script src="https://cdnjs.cloudflare.com/ajax/libs/three.js/r128/three.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/three@0.128.0/examples/js/loaders/GLTFLoader.js"></script>

<script>
const cfg = __CONFIG__;
const stage = document.getElementById("stage");
const statusText = document.getElementById("status");

if (typeof THREE === "undefined") {
    document.getElementById("fallback").style.display = "flex";
    statusText.textContent = "3D ENGINE UNAVAILABLE";
} else {

const scene = new THREE.Scene();
scene.fog = new THREE.Fog(0x061426, 8, 22);

const camera = new THREE.PerspectiveCamera(
    38, stage.clientWidth / stage.clientHeight, 0.1, 100
);
camera.position.set(0, 0.1, 8.5);

const renderer = new THREE.WebGLRenderer({
    antialias:true,
    alpha:true
});
renderer.setPixelRatio(Math.min(window.devicePixelRatio, 1.7));
renderer.setSize(stage.clientWidth, stage.clientHeight);
renderer.outputEncoding = THREE.sRGBEncoding;
stage.appendChild(renderer.domElement);

scene.add(new THREE.AmbientLight(0x9bcfff, 1.4));

const blueLight = new THREE.PointLight(0x008cff, 3, 12);
blueLight.position.set(-3, 3, 4);
scene.add(blueLight);

const goldLight = new THREE.PointLight(0xffbd54, 2, 10);
goldLight.position.set(3, 1, 3);
scene.add(goldLight);

const character = new THREE.Group();
scene.add(character);

const metal = new THREE.MeshStandardMaterial({
    color:0x9bb6ca,
    metalness:0.85,
    roughness:0.24
});

const skin = new THREE.MeshStandardMaterial({
    color:0xd5a98c,
    roughness:0.7
});

const dark = new THREE.MeshStandardMaterial({
    color:0x101d30,
    metalness:0.8,
    roughness:0.3
});

const blue = new THREE.MeshStandardMaterial({
    color:0x00aaff,
    emissive:0x0077ff,
    emissiveIntensity:2
});

const gold = new THREE.MeshStandardMaterial({
    color:0xffc45e,
    emissive:0x996000,
    emissiveIntensity:0.7,
    metalness:0.9
});

function sphere(parent, x,y,z, sx,sy,sz, material) {
    const mesh = new THREE.Mesh(
        new THREE.SphereGeometry(1,32,24),
        material
    );
    mesh.position.set(x,y,z);
    mesh.scale.set(sx,sy,sz);
    parent.add(mesh);
    return mesh;
}

function cylinder(parent, x,y,z, rt,rb,h, material) {
    const mesh = new THREE.Mesh(
        new THREE.CylinderGeometry(rt,rb,h,32),
        material
    );
    mesh.position.set(x,y,z);
    parent.add(mesh);
    return mesh;
}

// BODY
sphere(character,0,-1.25,0,0.95,1.05,0.48,dark);
cylinder(character,0,-0.48,0,0.30,0.34,0.55,metal);
sphere(character,0,-0.15,0,0.34,0.20,0.34,metal);

// CHEST ENERGY CORE
sphere(character,0,-1.12,0.47,0.19,0.19,0.08,blue);
sphere(character,0,-1.12,0.55,0.07,0.07,0.035,gold);

// SHOULDERS AND ARMS
for (const side of [-1,1]) {
    sphere(character,side*0.92,-0.52,0,0.32,0.34,0.34,metal);
    sphere(character,side*1.02,-1.05,0,0.20,0.55,0.22,dark);
    sphere(character,side*1.02,-1.55,0.04,0.22,0.23,0.23,metal);
}

// HEAD GROUP
const head = new THREE.Group();
head.position.set(0,0.78,0);
character.add(head);

sphere(head,0,0,0,0.69,0.88,0.57,skin);

// JAW
sphere(head,0,-0.36,0.12,0.48,0.38,0.43,skin);

// EARS
for (const side of [-1,1]) {
    sphere(head,side*0.68,-0.02,0,0.13,0.22,0.13,metal);
}

// ANDROID RIGHT TEMPLE
sphere(head,0.53,0.30,0.18,0.20,0.30,0.36,dark);
sphere(head,0.60,-0.10,0.27,0.16,0.22,0.19,metal);
sphere(head,0.62,0.34,0.36,0.07,0.07,0.05,blue);

// EYES
for (const side of [-1,1]) {
    sphere(head,side*0.27,0.12,0.49,0.15,0.075,0.055,dark);
    sphere(head,side*0.27,0.12,0.535,0.095,0.038,0.025,blue);
}

// NOSE
sphere(head,0,-0.02,0.57,0.09,0.17,0.10,skin);

// MOUTH
const mouth = sphere(
    head,0,-0.32,0.535,0.13,0.025,0.035,dark
);

// FOREHEAD LIGHT
sphere(head,0,0.65,0.35,0.08,0.08,0.04,gold);

// NECK TECHNOLOGY
sphere(character,0,-0.12,0.12,0.20,0.19,0.22,metal);

// ENERGY RINGS
const rings = [];
for (let i=0;i<3;i++) {
    const ring = new THREE.Mesh(
        new THREE.TorusGeometry(1.35+i*0.17,0.012,8,120),
        i===1 ? gold : blue
    );
    ring.rotation.x = Math.PI/2.7;
    ring.position.y = -0.6-i*0.18;
    character.add(ring);
    rings.push(ring);
}

// PARTICLES
const particleCount = 180;
const positions = new Float32Array(particleCount*3);

for(let i=0;i<particleCount;i++) {
    positions[i*3] = (Math.random()-0.5)*7;
    positions[i*3+1] = (Math.random()-0.5)*6;
    positions[i*3+2] = (Math.random()-0.5)*3;
}

const geometry = new THREE.BufferGeometry();
geometry.setAttribute(
    "position",new THREE.BufferAttribute(positions,3)
);

const particleMaterial = new THREE.PointsMaterial({
    color:0x35baff,
    size:0.025,
    transparent:true,
    opacity:0.8
});

scene.add(new THREE.Points(geometry,particleMaterial));

// OPTIONAL REAL 3D MODEL
if(cfg.modelUrl && THREE.GLTFLoader) {
    const loader = new THREE.GLTFLoader();
    loader.load(cfg.modelUrl,function(gltf) {
        const model = gltf.scene;
        model.scale.setScalar(1.4);
        model.position.y = -1.9;
        character.visible = false;
        scene.add(model);
        character.userData.realModel = model;
    },undefined,function() {
        character.visible = true;
    });
}

statusText.textContent =
    cfg.mode==="speaking" ? "JON IS SPEAKING" :
    cfg.mode==="listening" ? "JON IS LISTENING" :
    cfg.mode==="dance" ? "DANCE MODE" : "JON IS ONLINE";

const clock = new THREE.Clock();

function animate() {
    requestAnimationFrame(animate);
    const t = clock.getElapsedTime();

    // NATURAL MOVEMENT
    head.rotation.y = Math.sin(t*0.65)*0.09;
    head.rotation.x = Math.sin(t*0.8)*0.025;
    character.position.y = Math.sin(t*1.3)*0.06;

    // TALKING ANIMATION
    const talking = cfg.mode==="speaking";
    mouth.scale.y = talking ?
        0.025 + Math.abs(Math.sin(t*12))*0.13 : 0.025;

    // DANCE MODE
    if(cfg.mode==="dance") {
        character.rotation.y = Math.sin(t*2)*0.28;
        character.position.y = Math.abs(Math.sin(t*4))*0.22;
    } else {
        character.rotation.y = Math.sin(t*0.3)*0.04;
    }

    rings.forEach((ring,i)=>{
        ring.rotation.z = t*(i%2===0?0.25:-0.2);
    });

    renderer.render(scene,camera);
}

animate();

function resize() {
    const w=stage.clientWidth;
    const h=stage.clientHeight;
    camera.aspect=w/h;
    camera.updateProjectionMatrix();
    renderer.setSize(w,h);
}
window.addEventListener("resize",resize);

}
</script>
</div>
</body>
</html>
"""

    html = html.replace("HEIGHT", str(int(height)))
    html = html.replace("__CONFIG__", json.dumps(config))

    components.html(
        html,
        height=height,
        scrolling=False
    )
def show_character(height=620, speaking=False):
    animation = st.session_state.get(
        "character_animation", "idle"
    )

    if speaking:
        animation = "speaking"
    elif animation in ("Raqs", "Dance", "dance"):
        animation = "dance"
    elif animation in ("Tinglash", "listening"):
        animation = "listening"
    else:
        animation = "idle"

    render_character(
        height=height,
        name=st.session_state.get("character_name", "JON"),
        mode=animation,
        speaking=speaking,
        style=st.session_state.get(
            "character_style", "Human + Android"
        ),
        model_url=st.session_state.get(
            "character_model_url"
        )
    )
  
