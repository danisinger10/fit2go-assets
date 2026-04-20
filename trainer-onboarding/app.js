const state = {
  onboardingSteps: [
    {
      title: 'Week 1: Welcome & Systems Check',
      description:
        'Shadow senior coach, learn Fit2Go core values, and activate your personal dashboard.'
    },
    {
      title: 'Week 2: Client Experience Blueprint',
      description:
        'Run a mock consult, map client journey touchpoints, and commit to response time standards.'
    },
    {
      title: 'Week 3-4: Training Delivery Sprint',
      description:
        'Co-lead two virtual sessions, pilot our adaptive programming flow, and gather mentor feedback.'
    },
    {
      title: 'Week 5-8: Autonomy Accelerator',
      description:
        'Own 6 client programs, track metrics in the Fit2Go pulse sheet, and present insights to the team.'
    },
    {
      title: 'Week 9-12: Community Impact',
      description:
        'Launch a micro-community challenge, share success stories, and co-create new resources.'
    }
  ],
  microLessons: [
    {
      title: 'How Fit2Go Coaches Build Trust Fast',
      description:
        'Video breakdown of the first 15 minutes of a consult, with checklist for body language cues.'
    },
    {
      title: 'Adaptive Programming Cheatsheet',
      description: 'Downloadable flow chart that helps you adjust to client energy and schedule changes.'
    },
    {
      title: 'Feedback Framework',
      description: 'Micro-course on reflective listening and re-contracting goals when clients get stuck.'
    }
  ],
  focusAreas: [
    'Client Retention Rituals',
    'Adaptive Programming',
    'Data Storytelling',
    'Community Building',
    'Leadership & Mentoring'
  ],
  skills: [
    {
      name: 'Client Impact Storytelling',
      status: 'In Progress',
      action: 'Collect metrics & quotes from two clients',
      due: 'Apr 12'
    },
    {
      name: 'Adaptive Program Design',
      status: 'Needs Practice',
      action: 'Pair with mentor for 1-hour design jam',
      due: 'Apr 19'
    },
    {
      name: 'Community Activation',
      status: 'Ready to Launch',
      action: 'Schedule your first micro-challenge pilot',
      due: 'Apr 24'
    }
  ],
  radar: {
    labels: [
      'Client Experience',
      'Programming Mastery',
      'Communication',
      'Business Acumen',
      'Leadership',
      'Wellness Literacy'
    ],
    scores: [78, 72, 88, 64, 69, 80]
  },
  careerMilestones: [
    {
      title: '6 Month Marker: Lead Coach Certification',
      description:
        'Complete 100 coaching hours, deliver two signature workshops, and hit 90% client satisfaction.'
    },
    {
      title: '12 Month Marker: Regional Mentor',
      description:
        'Mentor two new hires, pilot a specialty track, and co-present at the Fit2Go virtual summit.'
    },
    {
      title: '18 Month Marker: Program Innovator',
      description:
        'Design a scalable habit challenge, increase retention by 12%, and publish a case study.'
    }
  ],
  highlights: [
    'Alicia just unlocked 95% client retention for Q1.',
    'New mobility toolkit dropped in the Resource Hub.',
    'Thursday huddle features guest mentor Tasha Greene.'
  ],
  experiments: [
    {
      title: 'Habit Momentum Pods',
      description: 'Test micro-accountability squads for remote clients.',
      tags: ['Community', 'Engagement', 'Beta'],
      completed: false
    },
    {
      title: 'Data Storytelling Looms',
      description: 'Record 3-minute weekly progress videos for leadership.',
      tags: ['Visibility', 'Leadership'],
      completed: false
    },
    {
      title: 'Partner Workout Exchange',
      description: 'Swap programming templates with another trainer to spark innovation.',
      tags: ['Collaboration'],
      completed: true
    }
  ],
  impactWins: [
    {
      title: 'Reduced churn by 18% in hybrid clients',
      description: 'Implemented tiered accountability nudges and Sunday prep sessions.',
      collaborators: ['Maya', 'Chris'],
      createdAt: 'Mar 21'
    }
  ]
};

const pulseRange = [82, 94];

const heroButtons = document.querySelectorAll('.hero__cta button');
const tabs = document.querySelectorAll('.tab-link');
const panels = document.querySelectorAll('.tab-panel');

const onboardingTimeline = document.getElementById('onboardingTimeline');
const microLearning = document.getElementById('microLearning');
const checkInForm = document.getElementById('checkInForm');
const checkInList = document.getElementById('checkInList');
const radarCanvas = document.getElementById('radarChart');
const radarLegend = document.getElementById('radarLegend');
const skillRows = document.getElementById('skillRows');
const impactForm = document.getElementById('impactForm');
const impactList = document.getElementById('impactList');
const experimentBoard = document.getElementById('experimentBoard');
const careerRoadmap = document.getElementById('careerRoadmap');
const weeklyHighlights = document.getElementById('weeklyHighlights');
const teamPulse = document.getElementById('teamPulse');

const highlightTemplate = document.getElementById('highlightTemplate');
const checkInTemplate = document.getElementById('checkInTemplate');
const skillRowTemplate = document.getElementById('skillRowTemplate');
const impactTemplate = document.getElementById('impactItemTemplate');
const experimentTemplate = document.getElementById('experimentTemplate');

const checkIns = [];

function renderHighlights() {
  weeklyHighlights.innerHTML = '';
  state.highlights.forEach((text) => {
    const li = highlightTemplate.content.cloneNode(true);
    li.querySelector('.text').textContent = text;
    weeklyHighlights.appendChild(li);
  });
}

function renderOnboardingTimeline() {
  onboardingTimeline.innerHTML = '';
  state.onboardingSteps.forEach((step) => {
    const li = document.createElement('li');
    const title = document.createElement('h3');
    const desc = document.createElement('p');
    title.textContent = step.title;
    desc.textContent = step.description;
    li.append(title, desc);
    onboardingTimeline.appendChild(li);
  });
}

function renderMicroLearning() {
  microLearning.innerHTML = '';
  state.microLessons.forEach((lesson) => {
    const article = document.createElement('article');
    const title = document.createElement('h3');
    const desc = document.createElement('p');
    title.textContent = lesson.title;
    desc.textContent = lesson.description;
    article.append(title, desc);
    microLearning.appendChild(article);
  });
}

function populateFocusAreas() {
  const select = checkInForm.querySelector('select[name="focus"]');
  state.focusAreas.forEach((area) => {
    const option = document.createElement('option');
    option.value = area;
    option.textContent = area;
    select.appendChild(option);
  });
}

function renderCheckIns() {
  checkInList.innerHTML = '';
  if (!checkIns.length) {
    const empty = document.createElement('li');
    empty.textContent = 'No check-ins scheduled yet. Add one to build momentum!';
    empty.classList.add('empty-state');
    checkInList.appendChild(empty);
    return;
  }

  checkIns.forEach((item, index) => {
    const node = checkInTemplate.content.cloneNode(true);
    node.querySelector('h4').textContent = item.focus;
    node.querySelector('.meta').textContent = `${item.date} · ${item.partner}`;
    const button = node.querySelector('button');
    button.addEventListener('click', () => {
      checkIns.splice(index, 1);
      renderCheckIns();
    });
    checkInList.appendChild(node);
  });
}

function renderSkills() {
  skillRows.innerHTML = '';
  state.skills.forEach((skill) => {
    const row = skillRowTemplate.content.cloneNode(true);
    row.querySelector('.skill').textContent = skill.name;
    row.querySelector('.status').textContent = skill.status;
    row.querySelector('.action').textContent = skill.action;
    row.querySelector('.due').textContent = skill.due;
    skillRows.appendChild(row);
  });
}

function renderImpactList() {
  impactList.innerHTML = '';
  state.impactWins.forEach((win) => {
    const item = impactTemplate.content.cloneNode(true);
    item.querySelector('h4').textContent = win.title;
    item.querySelector('.body').textContent = win.description;
    const collaborators = win.collaborators?.length ? ` · With ${win.collaborators.join(', ')}` : '';
    item.querySelector('.meta').textContent = `${win.createdAt}${collaborators}`;
    impactList.appendChild(item);
  });
}

function renderExperiments() {
  experimentBoard.innerHTML = '';
  state.experiments.forEach((experiment, index) => {
    const card = experimentTemplate.content.cloneNode(true);
    card.querySelector('h4').textContent = experiment.title;
    card.querySelector('.body').textContent = experiment.description;
    const tagWrapper = card.querySelector('.tags');

    experiment.tags.forEach((tag) => {
      const span = document.createElement('span');
      span.textContent = tag;
      tagWrapper.appendChild(span);
    });

    const button = card.querySelector('button');
    if (experiment.completed) {
      button.textContent = 'Completed';
      button.disabled = true;
      button.classList.add('completed');
    } else {
      button.addEventListener('click', () => {
        state.experiments[index].completed = true;
        renderExperiments();
      });
    }

    experimentBoard.appendChild(card);
  });
}

function renderCareerRoadmap() {
  careerRoadmap.innerHTML = '';
  state.careerMilestones.forEach((milestone) => {
    const wrapper = document.createElement('div');
    wrapper.className = 'milestone';
    const title = document.createElement('h3');
    const desc = document.createElement('p');
    title.textContent = milestone.title;
    desc.textContent = milestone.description;
    wrapper.append(title, desc);
    careerRoadmap.appendChild(wrapper);
  });
}

function renderRadarChart() {
  const ctx = radarCanvas.getContext('2d');
  const { labels, scores } = state.radar;
  const maxScore = 100;
  const totalPoints = labels.length;
  const center = { x: radarCanvas.width / 2, y: radarCanvas.height / 2 };
  const radius = Math.min(center.x, center.y) - 30;

  ctx.clearRect(0, 0, radarCanvas.width, radarCanvas.height);
  ctx.strokeStyle = 'rgba(19, 37, 52, 0.25)';
  ctx.lineWidth = 1;

  const levels = 5;
  for (let level = 1; level <= levels; level++) {
    const levelRadius = (radius * level) / levels;
    ctx.beginPath();
    for (let i = 0; i < totalPoints; i++) {
      const angle = (Math.PI * 2 * i) / totalPoints;
      const x = center.x + levelRadius * Math.sin(angle);
      const y = center.y - levelRadius * Math.cos(angle);
      if (i === 0) {
        ctx.moveTo(x, y);
      } else {
        ctx.lineTo(x, y);
      }
    }
    ctx.closePath();
    ctx.stroke();
  }

  ctx.beginPath();
  for (let i = 0; i < totalPoints; i++) {
    const angle = (Math.PI * 2 * i) / totalPoints;
    const scoreRadius = (radius * scores[i]) / maxScore;
    const x = center.x + scoreRadius * Math.sin(angle);
    const y = center.y - scoreRadius * Math.cos(angle);
    if (i === 0) {
      ctx.moveTo(x, y);
    } else {
      ctx.lineTo(x, y);
    }
  }
  ctx.closePath();
  ctx.fillStyle = 'rgba(57, 194, 127, 0.35)';
  ctx.fill();
  ctx.strokeStyle = 'rgba(33, 155, 93, 0.8)';
  ctx.stroke();

  ctx.fillStyle = 'rgba(19, 37, 52, 0.65)';
  ctx.font = '13px Manrope, sans-serif';
  labels.forEach((label, i) => {
    const angle = (Math.PI * 2 * i) / totalPoints;
    const labelRadius = radius + 18;
    const x = center.x + labelRadius * Math.sin(angle);
    const y = center.y - labelRadius * Math.cos(angle);
    ctx.textAlign = 'center';
    ctx.fillText(label, x, y);
  });

  radarLegend.innerHTML = '';
  const current = document.createElement('span');
  current.innerHTML = '<span class="dot" style="background: rgba(57, 194, 127, 0.7);"></span> Current Strength';
  radarLegend.appendChild(current);
}

function animatePulse() {
  const [min, max] = pulseRange;
  const next = Math.floor(Math.random() * (max - min + 1) + min);
  teamPulse.textContent = `${next}%`;
}

function bindEvents() {
  heroButtons.forEach((btn) =>
    btn.addEventListener('click', () => {
      const tabId = btn.dataset.target;
      switchTab(tabId);
      window.scrollTo({ top: document.querySelector('main').offsetTop - 40, behavior: 'smooth' });
    })
  );

  tabs.forEach((tab) =>
    tab.addEventListener('click', () => {
      switchTab(tab.dataset.tab);
    })
  );

  checkInForm.addEventListener('submit', (event) => {
    event.preventDefault();
    const formData = new FormData(checkInForm);
    const focus = formData.get('focus');
    const date = formData.get('date');
    const partner = formData.get('partner');

    if (!focus || !date || !partner) {
      return;
    }

    checkIns.push({ focus, date, partner });
    renderCheckIns();
    checkInForm.reset();
  });

  impactForm.addEventListener('submit', (event) => {
    event.preventDefault();
    const formData = new FormData(impactForm);
    const title = formData.get('title');
    const description = formData.get('description');
    const collaborators = formData
      .get('collaborators')
      .split(',')
      .map((name) => name.trim())
      .filter(Boolean);

    state.impactWins.unshift({
      title,
      description,
      collaborators,
      createdAt: new Date().toLocaleDateString('en-US', {
        month: 'short',
        day: 'numeric'
      })
    });

    renderImpactList();
    impactForm.reset();
  });
}

function switchTab(tabId) {
  tabs.forEach((tab) => tab.classList.toggle('active', tab.dataset.tab === tabId));
  panels.forEach((panel) => panel.classList.toggle('active', panel.id === tabId));
}

function init() {
  renderHighlights();
  renderOnboardingTimeline();
  renderMicroLearning();
  populateFocusAreas();
  renderCheckIns();
  renderSkills();
  renderImpactList();
  renderExperiments();
  renderCareerRoadmap();
  renderRadarChart();
  bindEvents();
  animatePulse();
  setInterval(animatePulse, 7000);
}

init();
