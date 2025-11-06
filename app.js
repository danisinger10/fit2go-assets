const onboardingTasks = [
  {
    id: 'values',
    title: 'Fit2Go values immersion',
    description: 'Watch the founder story, then share your takeaways in the cohort channel.',
    duration: 'Day 1',
  },
  {
    id: 'tech',
    title: 'Platform tech check',
    description: 'Complete the live systems walkthrough and set up your scheduling + billing automations.',
    duration: 'Day 2',
  },
  {
    id: 'shadow',
    title: 'Shadow a live session',
    description: 'Observe a senior coach and log three coaching moments you want to replicate.',
    duration: 'Day 4',
  },
  {
    id: 'practice',
    title: 'Deliver your practice session',
    description: 'Run a 30-minute mock session with mentor feedback focused on cueing and pacing.',
    duration: 'Day 7',
  },
  {
    id: 'clients',
    title: 'Activate first client roster',
    description: 'Set 30-day goals with your first five clients and load plans into Trainer HQ.',
    duration: 'Day 10',
  },
  {
    id: 'retro',
    title: 'Onboarding retrospective',
    description: 'Meet with your mentor to review KPIs and commit to your next mastery sprint.',
    duration: 'Day 14',
  },
];

const milestones = [
  {
    title: 'Certification & compliance',
    detail: 'Submit CPR, liability waiver, and Fit2Go hybrid training certification.',
    meta: 'Complete by Day 3',
  },
  {
    title: 'Coach performance scorecard',
    detail: 'Score 80%+ on empathy, cueing, and accountability during mentor review.',
    meta: 'Complete by Day 9',
  },
  {
    title: 'Client satisfaction baseline',
    detail: 'Achieve 4.6+ post-session rating and 90% habit streak across first five clients.',
    meta: 'Complete by Day 14',
  },
];

const focusWeeks = [
  {
    title: 'Connect & calibrate',
    description: 'Deep dive on Fit2Go culture, meet your mentor, and align expectations.',
    meta: 'Week 1',
  },
  {
    title: 'Coach with confidence',
    description: 'Deliver two mock sessions, scorecards, and calibrate your cueing system.',
    meta: 'Week 2',
  },
  {
    title: 'Launch client impact',
    description: 'Activate first roster, review engagement data, and plan your mastery sprint.',
    meta: 'Week 3',
  },
];

const skillTracks = [
  {
    id: 'activeListening',
    category: 'coaching',
    title: 'Motivational interviewing & active listening',
    level: 'Level 2 mastery',
    focus: ['Mirroring', 'Powerful questioning', 'Confidence calibration'],
    sprint: 'Attend MI lab + submit 2 recorded sessions',
    badge: 'Coaching excellence',
  },
  {
    id: 'cueing',
    category: 'coaching',
    title: 'Precision cueing for hybrid sessions',
    level: 'Level 1 mastery',
    focus: ['Virtual + in-person cues', 'Corrective exercise flow'],
    sprint: 'Shadow Elite Coach + deliver hybrid workshop',
    badge: 'Hybrid coaching',
  },
  {
    id: 'programming',
    category: 'programming',
    title: 'Habit-first program design',
    level: 'Level 3 mastery',
    focus: ['Habit stacking', 'Recovery protocols', 'Progressive overload scripts'],
    sprint: 'Build 3 habit-first plans + peer review',
    badge: 'Program design',
  },
  {
    id: 'dataFluency',
    category: 'business',
    title: 'Client insights & retention analytics',
    level: 'Level 2 mastery',
    focus: ['Engagement dashboards', 'Retention levers'],
    sprint: 'Partner with Ops on churn deep dive',
    badge: 'Client success',
  },
  {
    id: 'sales',
    category: 'business',
    title: 'Consultative sales for warm leads',
    level: 'Level 1 mastery',
    focus: ['Discovery frameworks', 'Objection handling'],
    sprint: 'Complete Sales Studio micro-cohort',
    badge: 'Revenue impact',
  },
  {
    id: 'systems',
    category: 'programming',
    title: 'Automation & systems playbook',
    level: 'Level 2 mastery',
    focus: ['Zapier flows', 'Template optimization'],
    sprint: 'Ship 2 automations w/ Ops partner',
    badge: 'Operational excellence',
  },
];

const careerPaths = {
  eliteCoach: {
    title: 'Elite Coach',
    summary: 'Deliver signature Fit2Go results, lead client intensives, and mentor new trainers.',
    metrics: ['CSAT 4.8+', 'Client retention 93%', 'Referral growth +20%'],
    experiences: ['Lead Fit2Go Challenge cohort', 'Publish 2 expert workshops', 'Coach mentor rotation'],
    enablers: ['Advanced cueing lab', 'Specialty certification stipend', 'Mentor feedback loops'],
  },
  teamLead: {
    title: 'Regional Team Lead',
    summary: 'Drive performance for 6-8 coaches, scale best practices, and deliver regional workshops.',
    metrics: ['Team NPS 60+', 'Regional revenue +18%', 'Coach ramp time 25 days'],
    experiences: ['Lead quarterly team summit', 'Build hiring scorecard', 'Launch retention experiments'],
    enablers: ['Leadership residency', 'Data storytelling clinic', 'Ops partnership playbook'],
  },
  studioDirector: {
    title: 'Studio Director',
    summary: 'Own P&L, launch new locations, and partner with Growth to expand the Fit2Go footprint.',
    metrics: ['Studio EBITDA 22%', 'Launch 2 satellite pods', 'Headcount engagement 85%'],
    experiences: ['Design market launch plan', 'Co-lead investor update', 'Coach succession design'],
    enablers: ['Executive mentor', 'Strategic finance sprints', 'Growth experimentation budget'],
  },
};

const resourceLibrary = [
  {
    title: 'Trainer HQ quickstart',
    type: 'Playbook',
    duration: '25 min read',
    description: 'Master scheduling, billing, and retention automations in the Fit2Go platform.',
  },
  {
    title: 'Client storytelling templates',
    type: 'Template pack',
    duration: '10 files',
    description: 'Share progress with clients using customizable milestone and habit dashboards.',
  },
  {
    title: 'Hybrid coaching best practices',
    type: 'Workshop replay',
    duration: '48 min',
    description: 'Blend in-person energy with remote accountability for modern professionals.',
  },
  {
    title: 'Leadership fast track',
    type: 'Live cohort',
    duration: 'Next start: May 6',
    description: 'Build coaching teams, manage performance, and scale Fit2Go experiences.',
  },
];

const completedTasks = new Set();

function renderOnboardingTasks() {
  const container = document.getElementById('onboarding-tasks');
  container.innerHTML = '';

  onboardingTasks.forEach((task) => {
    const li = document.createElement('li');
    li.innerHTML = `
      <div class="checklist__header">
        <input type="checkbox" id="task-${task.id}" data-task="${task.id}" ${
          completedTasks.has(task.id) ? 'checked' : ''
        } />
        <label for="task-${task.id}">${task.title}</label>
        <span class="badge">${task.duration}</span>
      </div>
      <p class="checklist__description">${task.description}</p>
    `;
    container.appendChild(li);
  });
}

function renderMilestones() {
  const container = document.getElementById('onboarding-milestones');
  container.innerHTML = '';

  milestones.forEach((milestone) => {
    const item = document.createElement('li');
    item.className = 'milestone';
    item.innerHTML = `
      <div class="timeline__title">${milestone.title}</div>
      <div>${milestone.detail}</div>
      <div class="milestone__meta">${milestone.meta}</div>
    `;
    container.appendChild(item);
  });
}

function renderFocusWeeks() {
  const container = document.getElementById('weekly-focus');
  focusWeeks.forEach((week) => {
    const item = document.createElement('li');
    item.innerHTML = `
      <span class="timeline__title">${week.title}</span>
      <span class="timeline__meta">${week.meta}</span>
      <p class="timeline__description">${week.description}</p>
    `;
    container.appendChild(item);
  });
}

function updateProgress() {
  const progressBar = document.getElementById('onboarding-progress');
  const label = document.getElementById('onboarding-progress-label');
  const percent = Math.round((completedTasks.size / onboardingTasks.length) * 100);
  progressBar.style.width = `${percent}%`;
  label.textContent = `${percent}% complete`;
}

function handleTaskClick(event) {
  const checkbox = event.target.closest('input[type="checkbox"]');
  if (!checkbox) return;
  const { task } = checkbox.dataset;

  if (checkbox.checked) {
    completedTasks.add(task);
  } else {
    completedTasks.delete(task);
  }

  updateProgress();
}

function renderSkillCards(filter = 'all') {
  const container = document.getElementById('skill-cards');
  container.innerHTML = '';

  skillTracks
    .filter((track) => filter === 'all' || track.category === filter)
    .forEach((track) => {
      const card = document.createElement('article');
      card.className = 'skill-card';
      card.innerHTML = `
        <div class="skill-card__header">
          <h3>${track.title}</h3>
          <span class="skill-card__tag">${track.badge}</span>
        </div>
        <div class="skill-card__level">${track.level}</div>
        <div>
          <h4>Focus reps</h4>
          <ul>${track.focus.map((item) => `<li>${item}</li>`).join('')}</ul>
        </div>
        <div class="skill-card__actions">
          <span class="badge">Sprint: ${track.sprint}</span>
          <button class="btn btn--ghost" data-track="${track.id}">Add to sprint</button>
        </div>
      `;
      container.appendChild(card);
    });
}

function updateFilterButtons(activeFilter) {
  document.querySelectorAll('.filter-btn').forEach((button) => {
    button.classList.toggle('is-active', button.dataset.filter === activeFilter);
  });
}

function renderCareerCard(pathKey, timelineMonths) {
  const container = document.getElementById('career-details');
  const data = careerPaths[pathKey];
  const paceDescriptor = timelineMonths <= 9 ? 'Accelerated' : timelineMonths >= 18 ? 'Deliberate' : 'Standard';

  container.className = 'card career-card';
  container.innerHTML = `
    <div class="career-card__header">
      <h3>${data.title}</h3>
      <p class="career-card__meta">${data.summary}</p>
      <div class="pill pill--info">${paceDescriptor} pace · ${timelineMonths} month plan</div>
    </div>
    <div class="career-card__grid">
      <section>
        <h4>Key metrics</h4>
        <ul>${data.metrics.map((metric) => `<li>${metric}</li>`).join('')}</ul>
      </section>
      <section>
        <h4>Signature experiences</h4>
        <ul>${data.experiences.map((experience) => `<li>${experience}</li>`).join('')}</ul>
      </section>
      <section>
        <h4>Enablement & support</h4>
        <ul>${data.enablers.map((enabler) => `<li>${enabler}</li>`).join('')}</ul>
      </section>
    </div>
  `;
}

function renderResources() {
  const container = document.getElementById('resource-grid');
  resourceLibrary.forEach((resource) => {
    const card = document.createElement('article');
    card.className = 'resource-card';
    card.innerHTML = `
      <h3>${resource.title}</h3>
      <div class="resource-card__meta">
        <span>${resource.type}</span>
        <span>${resource.duration}</span>
      </div>
      <p>${resource.description}</p>
      <button class="btn btn--ghost">Save to my library</button>
    `;
    container.appendChild(card);
  });
}

function setupInteractions() {
  document.getElementById('onboarding-tasks').addEventListener('click', handleTaskClick);

  document.querySelectorAll('.filter-btn').forEach((button) => {
    button.addEventListener('click', () => {
      const { filter } = button.dataset;
      updateFilterButtons(filter);
      renderSkillCards(filter);
    });
  });

  const careerSelect = document.getElementById('career-select');
  const timelineSlider = document.getElementById('timeline-slider');
  const timelineValue = document.getElementById('timeline-value');

  function updateCareerCard() {
    renderCareerCard(careerSelect.value, Number(timelineSlider.value));
  }

  careerSelect.addEventListener('change', updateCareerCard);
  timelineSlider.addEventListener('input', () => {
    timelineValue.textContent = `${timelineSlider.value} months`;
    updateCareerCard();
  });

  document.getElementById('tour-app').addEventListener('click', () => {
    document.getElementById('tour-dialog').showModal();
  });

  document.getElementById('download-plan').addEventListener('click', () => {
    document.getElementById('plan-dialog').showModal();
  });

  document.getElementById('add-to-plan').addEventListener('click', () => {
    const planDialog = document.getElementById('plan-dialog');
    if (!planDialog.open) {
      planDialog.showModal();
    }
  });

  document.getElementById('start-journey').addEventListener('click', () => {
    document.getElementById('onboarding').scrollIntoView({ behavior: 'smooth' });
  });

  document.getElementById('share-progress').addEventListener('click', () => {
    navigator.clipboard
      .writeText('Check out my Fit2Go development plan!')
      .then(() => {
        const footerButton = document.getElementById('share-progress');
        footerButton.textContent = 'Copied to clipboard!';
        setTimeout(() => {
          footerButton.textContent = 'Share progress';
        }, 2000);
      })
      .catch(() => {
        alert('Copy failed. Try again or share manually.');
      });
  });
}

function init() {
  renderOnboardingTasks();
  renderMilestones();
  renderFocusWeeks();
  renderSkillCards();
  renderCareerCard('eliteCoach', Number(document.getElementById('timeline-slider').value));
  renderResources();
  updateProgress();
  setupInteractions();
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', init);
} else {
  init();
}
