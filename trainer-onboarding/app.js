const onboardingMilestones = [
  {
    id: "foundation",
    title: "Week 0: Welcome & Foundations",
    badge: "Start",
    description:
      "Meet your mentor, align on Fit2Go's values, and complete the orientation playlist so you know what excellence looks like.",
    checklist: [
      {
        title: "Kickoff huddle",
        description: "Meet your mentor and outline your 2-week onboarding plan.",
        duration: "45 min",
        owner: "Mentor",
      },
      {
        title: "Culture immersion",
        description: "Complete the Fit2Go story, manifesto, and service standards modules.",
        duration: "90 min",
        owner: "You",
      },
      {
        title: "Systems setup",
        description: "Configure coaching dashboard, client messaging, and weekly reporting tools.",
        duration: "60 min",
        owner: "Ops",
      },
    ],
    stats: { hours: 3, skills: 2, touchpoints: 2 },
  },
  {
    id: "skills",
    title: "Week 1: Client Delivery Skills",
    badge: "Build",
    description:
      "Shadow live sessions, practice key service frameworks, and run your first mock client consult.",
    checklist: [
      {
        title: "Shadow 3 sessions",
        description: "Observe veteran trainers and capture your takeaways in the reflection doc.",
        duration: "3 hrs",
        owner: "You",
      },
      {
        title: "Service framework lab",
        description: "Practice onboarding, progression, and accountability scripts with your mentor.",
        duration: "2 hrs",
        owner: "Mentor",
      },
      {
        title: "Mock consult",
        description: "Run a simulated session and earn your green-light certification.",
        duration: "75 min",
        owner: "QA Lead",
      },
    ],
    stats: { hours: 5.25, skills: 4, touchpoints: 3 },
  },
  {
    id: "launch",
    title: "Week 2: Client Launch",
    badge: "Launch",
    description:
      "Activate your first Fit2Go client pod, deliver onboarding communications, and align on success metrics.",
    checklist: [
      {
        title: "Client pod assignment",
        description: "Review client dossiers and develop 30-day training roadmaps.",
        duration: "2 hrs",
        owner: "Ops",
      },
      {
        title: "Accountability cadence",
        description: "Schedule weekly coaching calls and automation touchpoints.",
        duration: "45 min",
        owner: "You",
      },
      {
        title: "Launch retro",
        description: "Meet with mentor to review launch metrics and growth opportunities.",
        duration: "60 min",
        owner: "Mentor",
      },
    ],
    stats: { hours: 3.75, skills: 3, touchpoints: 4 },
  },
];

const skillTracks = [
  {
    id: "program-design",
    title: "Program Design & Adaptation",
    description:
      "Craft bespoke training blocks, leverage Fit2Go's data engine, and iterate based on client feedback.",
    progress: 60,
    focusAreas: ["Creative block design", "Behavior-driven adjustments", "Recovery protocols"],
    actions: ["Book a programming lab", "Download template", "Request mentor feedback"],
  },
  {
    id: "coaching",
    title: "Coaching Communication",
    description:
      "Master motivational interviewing, habit coaching, and accountability messaging that keeps clients engaged.",
    progress: 45,
    focusAreas: ["Motivational interviewing", "Habit stacking", "Automation scripts"],
    actions: ["Schedule shadow session", "Practice objection handling", "Review best practices"],
  },
  {
    id: "data",
    title: "Performance Analytics",
    description:
      "Interpret readiness scores, translate data into training pivots, and present insights during client reviews.",
    progress: 30,
    focusAreas: ["Readiness dashboards", "Insight storytelling", "Quarterly reviews"],
    actions: ["Complete analytics micro-lesson", "Submit case study", "Join office hours"],
  },
  {
    id: "business",
    title: "Business Growth",
    description:
      "Build referral engines, nurture community events, and unlock leadership tracks within Fit2Go.",
    progress: 20,
    focusAreas: ["Referral systems", "Workshop facilitation", "Mentor leadership"],
    actions: ["Pitch a partner event", "Shadow a sales consult", "Apply for squad lead"],
  },
];

const careerPaths = [
  {
    id: "specialist",
    title: "Elite Client Specialist",
    summary: "Deepen mastery with high-touch executive clients and recovery protocols.",
    roadmap: [
      {
        title: "Quarter 1",
        items: [
          "Earn Mobility & Recovery Specialist credential",
          "Shadow 2 executive client re-assessments",
          "Lead a recovery-focused community session",
        ],
      },
      {
        title: "Quarter 2",
        items: [
          "Launch advanced biometric tracking pilot",
          "Co-create 3 executive client case studies",
          "Mentor a new hire on precision programming",
        ],
      },
      {
        title: "Quarter 3",
        items: [
          "Present insights during leadership sync",
          "Pilot concierge coaching enhancements",
          "Evaluate ROI for recovery service bundle",
        ],
      },
    ],
  },
  {
    id: "leader",
    title: "Squad Leader",
    summary: "Scale impact by leading a pod of trainers and driving quality standards.",
    roadmap: [
      {
        title: "Quarter 1",
        items: [
          "Complete leadership foundations accelerator",
          "Run weekly pod retros with QA partner",
          "Own onboarding for 1 new hire",
        ],
      },
      {
        title: "Quarter 2",
        items: [
          "Certify in coaching QA rubric",
          "Launch trainer scorecard reviews",
          "Host monthly team building experience",
        ],
      },
      {
        title: "Quarter 3",
        items: [
          "Co-lead hiring panel for next cohort",
          "Roll out advanced client retention playbook",
          "Pilot new accountability automation",
        ],
      },
    ],
  },
  {
    id: "entrepreneur",
    title: "Market Expansion Entrepreneur",
    summary: "Drive new market launches, partnerships, and Fit2Go brand presence.",
    roadmap: [
      {
        title: "Quarter 1",
        items: [
          "Develop go-to-market brief for new territory",
          "Secure 3 corporate wellness presentations",
          "Collaborate with marketing on local playbook",
        ],
      },
      {
        title: "Quarter 2",
        items: [
          "Pilot micro-gym partnership strategy",
          "Design pop-up event series",
          "Report on acquisition funnel performance",
        ],
      },
      {
        title: "Quarter 3",
        items: [
          "Launch ambassador program",
          "Negotiate co-branded wellness experiences",
          "Scale operations hand-off plan",
        ],
      },
    ],
  },
];

const resources = [
  {
    title: "First 30 Days Playbook",
    description: "Daily checkpoints, templates, and outcomes for your ramp period.",
    highlights: [
      "Two-week onboarding checklist",
      "Client pod communication scripts",
      "Feedback loop prompts",
    ],
  },
  {
    title: "Consult Blueprint",
    description: "Conversation guides and frameworks to run high-converting consults.",
    highlights: [
      "Discovery questions",
      "Objection handling matrix",
      "Call flow examples",
    ],
  },
  {
    title: "Program Library",
    description: "Ready-to-run training blocks for popular goals and client segments.",
    highlights: [
      "Strength & power microcycles",
      "Executive travel adaptations",
      "Low back rehab progressions",
    ],
  },
  {
    title: "Growth Engine Toolkit",
    description: "Referrals, partnerships, and community event templates for expansion.",
    highlights: [
      "Email & SMS nurture sequences",
      "Event blueprint",
      "Partnership scorecard",
    ],
  },
];

const stats = onboardingMilestones.reduce(
  (acc, milestone) => {
    acc.hours += milestone.stats.hours;
    acc.skills += milestone.stats.skills;
    acc.touchpoints += milestone.stats.touchpoints;
    return acc;
  },
  { hours: 0, skills: 0, touchpoints: 0 }
);

const heroStats = {
  hours: document.getElementById("total-hours"),
  skills: document.getElementById("core-skills"),
  touchpoints: document.getElementById("mentorship-touchpoints"),
};

const onboardingTimelineEl = document.getElementById("onboardingTimeline");
const checklistEl = document.getElementById("checklist");
const focusChipsEl = document.getElementById("focusChips");
const skillGridEl = document.getElementById("skillGrid");
const pathSelectorEl = document.getElementById("pathSelector");
const roadmapEl = document.getElementById("roadmap");
const resourceGridEl = document.getElementById("resourceGrid");
const resourceModalEl = document.getElementById("resourceModal");
const resourceTitleEl = document.getElementById("resourceTitle");
const resourceDescriptionEl = document.getElementById("resourceDescription");
const resourceHighlightsEl = document.getElementById("resourceHighlights");
const closeModalBtn = resourceModalEl.querySelector(".close");

let activeMilestone = onboardingMilestones[0].id;
let activePath = careerPaths[0].id;
let activeChipIndex = 0;

function animateNumber(element, target, duration = 900) {
  const start = performance.now();
  const initial = 0;
  const update = (time) => {
    const progress = Math.min((time - start) / duration, 1);
    const value = Math.floor(initial + progress * target);
    element.textContent = value;
    if (progress < 1) requestAnimationFrame(update);
  };
  requestAnimationFrame(update);
}

function renderStats() {
  animateNumber(heroStats.hours, Math.round(stats.hours));
  animateNumber(heroStats.skills, stats.skills);
  animateNumber(heroStats.touchpoints, stats.touchpoints);
}

function renderTimeline() {
  onboardingTimelineEl.innerHTML = onboardingMilestones
    .map((milestone) => {
      const activeClass = milestone.id === activeMilestone ? "active" : "";
      return `
        <article class="timeline-card ${activeClass}" data-id="${milestone.id}">
          <span class="badge">${milestone.badge}</span>
          <h3>${milestone.title}</h3>
          <p>${milestone.description}</p>
        </article>
      `;
    })
    .join("");
}

function renderChecklist() {
  const milestone = onboardingMilestones.find((item) => item.id === activeMilestone);
  checklistEl.innerHTML = milestone.checklist
    .map(
      (task) => `
        <article class="task-card">
          <header>
            <h3>${task.title}</h3>
            <span class="badge">${task.owner}</span>
          </header>
          <p>${task.description}</p>
          <div class="meta">
            <span>${task.duration}</span>
            <button type="button" data-title="${task.title}">Mark Complete</button>
          </div>
        </article>
      `
    )
    .join("");
}

function renderFocusChips() {
  const chips = ["All", ...skillTracks.map((track) => track.title)];
  focusChipsEl.innerHTML = chips
    .map(
      (label, index) => `
        <button class="chip ${index === activeChipIndex ? "active" : ""}" data-index="${index}">${label}</button>
      `
    )
    .join("");
}

function renderSkills() {
  let filteredTracks = skillTracks;
  if (activeChipIndex > 0) {
    const selectedTitle = skillTracks[activeChipIndex - 1].title;
    filteredTracks = skillTracks.filter((track) => track.title === selectedTitle);
  }

  skillGridEl.innerHTML = filteredTracks
    .map(
      (track) => `
        <article class="skill-card">
          <header>
            <h3>${track.title}</h3>
            <div class="progress"><span style="width:${track.progress}%"></span></div>
          </header>
          <p>${track.description}</p>
          <ul>
            ${track.focusAreas.map((focus) => `<li>${focus}</li>`).join("")}
          </ul>
          <div class="skill-actions">
            ${track.actions
              .map((action) => `<button type="button" data-action="${action}">${action}</button>`)
              .join("")}
          </div>
        </article>
      `
    )
    .join("");
}

function renderPaths() {
  pathSelectorEl.innerHTML = careerPaths
    .map(
      (path) => `
        <article class="path-card ${path.id === activePath ? "active" : ""}" data-id="${path.id}">
          <h3>${path.title}</h3>
          <p>${path.summary}</p>
        </article>
      `
    )
    .join("");
}

function renderRoadmap() {
  const path = careerPaths.find((item) => item.id === activePath);
  roadmapEl.innerHTML = path.roadmap
    .map(
      (phase) => `
        <article class="roadmap-card">
          <h4>${phase.title}</h4>
          <ul>${phase.items.map((item) => `<li>${item}</li>`).join("")}</ul>
        </article>
      `
    )
    .join("");
}

function renderResources() {
  resourceGridEl.innerHTML = resources
    .map(
      (resource, index) => `
        <article class="resource-card">
          <h3>${resource.title}</h3>
          <p>${resource.description}</p>
          <button type="button" data-index="${index}">View Details</button>
        </article>
      `
    )
    .join("");
}

function openResourceModal(resource) {
  resourceTitleEl.textContent = resource.title;
  resourceDescriptionEl.textContent = resource.description;
  resourceHighlightsEl.innerHTML = resource.highlights.map((item) => `<li>${item}</li>`).join("");
  resourceModalEl.classList.add("open");
  resourceModalEl.setAttribute("aria-hidden", "false");
  resourceModalEl.setAttribute("aria-modal", "true");
}

function closeResourceModal() {
  resourceModalEl.classList.remove("open");
  resourceModalEl.setAttribute("aria-hidden", "true");
  resourceModalEl.setAttribute("aria-modal", "false");
}

function handleTimelineClick(event) {
  const card = event.target.closest(".timeline-card");
  if (!card) return;
  activeMilestone = card.dataset.id;
  renderTimeline();
  renderChecklist();
}

function handleChecklistClick(event) {
  const button = event.target.closest("button[data-title]");
  if (!button) return;
  button.classList.toggle("completed");
  button.textContent = button.classList.contains("completed") ? "Completed" : "Mark Complete";
}

function handleChipClick(event) {
  const chip = event.target.closest(".chip");
  if (!chip) return;
  activeChipIndex = Number(chip.dataset.index);
  focusChipsEl.querySelectorAll(".chip").forEach((node) => node.classList.remove("active"));
  chip.classList.add("active");
  renderSkills();
}

function handlePathSelect(event) {
  const card = event.target.closest(".path-card");
  if (!card) return;
  activePath = card.dataset.id;
  renderPaths();
  renderRoadmap();
}

function handleResourceClick(event) {
  const button = event.target.closest("button[data-index]");
  if (!button) return;
  const resource = resources[Number(button.dataset.index)];
  openResourceModal(resource);
}

function handleModalClick(event) {
  if (event.target === resourceModalEl) {
    closeResourceModal();
  }
}

function handleScrollAction(event) {
  const button = event.target.closest("[data-scroll]");
  if (!button) return;
  const target = document.querySelector(button.dataset.scroll);
  if (target) {
    target.scrollIntoView({ behavior: "smooth" });
  }
}

function handleKeydown(event) {
  if (event.key === "Escape" && resourceModalEl.classList.contains("open")) {
    closeResourceModal();
  }
}

renderStats();
renderTimeline();
renderChecklist();
renderFocusChips();
renderSkills();
renderPaths();
renderRoadmap();
renderResources();

onboardingTimelineEl.addEventListener("click", handleTimelineClick);
checklistEl.addEventListener("click", handleChecklistClick);
focusChipsEl.addEventListener("click", handleChipClick);
pathSelectorEl.addEventListener("click", handlePathSelect);
resourceGridEl.addEventListener("click", handleResourceClick);
resourceModalEl.addEventListener("click", handleModalClick);
closeModalBtn.addEventListener("click", closeResourceModal);
document.body.addEventListener("click", handleScrollAction);
document.addEventListener("keydown", handleKeydown);
