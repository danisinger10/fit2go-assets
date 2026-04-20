const onboardingSteps = [
  {
    title: 'Orientation & Fit2Go DNA',
    summary:
      'Immerse yourself in the brand story, service promise, and our trainer community. Finish the day by booking your welcome coaching session.',
    tasks: [
      {
        label: 'Watch the founder welcome video and note two takeaways to share in the community channel.',
        detail: 'Access via the Fit2Go Academy portal.',
      },
      {
        label: 'Complete the “Who I Am as a Coach” profile for the Fit2Go directory.',
        detail: 'Upload your headshot, bio, and signature training win.',
      },
      {
        label: 'Schedule your mentor matching call.',
        detail: 'Introduce yourself and identify areas you want to grow first.',
      },
    ],
  },
  {
    title: 'Operational Foundations',
    summary:
      'Build confidence in Fit2Go systems and client touchpoints. Deliver your first practice session with feedback.',
    tasks: [
      {
        label: 'Pass the Fit2Go tech walkthrough.',
        detail: 'Test scheduling, session notes, and KPI dashboards with a teammate.',
      },
      {
        label: 'Shadow a senior trainer session.',
        detail: 'Capture three coaching moves that match the Fit2Go methodology.',
      },
      {
        label: 'Deliver a mock client session.',
        detail: 'Record yourself, share with your mentor, and review the session rubric.',
      },
    ],
  },
  {
    title: 'Client Delivery Mastery',
    summary:
      'Showcase consistent client outcomes and refine your personal coaching brand.',
    tasks: [
      {
        label: 'Launch a 4-week transformation plan with two clients.',
        detail: 'Submit kickoff outlines to your mentor within 48 hours of first session.',
      },
      {
        label: 'Complete the behavior change micro-course.',
        detail: 'Score 85% or above on the assessment to unlock live lab access.',
      },
      {
        label: 'Host a Fit2Go community workshop.',
        detail: 'Pick a topic aligned with your strengths and collect participant feedback.',
      },
    ],
  },
  {
    title: 'Growth & Ownership',
    summary:
      'Create momentum for long-term success. Align with Fit2Go leadership on performance metrics and new opportunities.',
    tasks: [
      {
        label: 'Present your 90-day impact report.',
        detail: 'Highlight KPI growth, testimonials, and next quarter focus.',
      },
      {
        label: 'Design a signature specialty offer.',
        detail: 'Draft pricing, outcome promise, and marketing hooks.',
      },
      {
        label: 'Book a quarterly career planning session.',
        detail: 'Align on advancement pathways and new responsibilities.',
      },
    ],
  },
];

const skillAreas = [
  {
    name: 'Client Acquisition & Relationship Building',
    description:
      'Cultivate referral engines, host irresistible consultations, and create retention through white-glove service.',
    trackerCaption: 'Log each outreach sprint and celebrate every new consultation booked.',
    tasks: [
      'Map 5 high-potential referral partners and send introduction messages.',
      'Host two “movement audit” consultations with feedback from sales enablement.',
      'Implement a two-touch post-session follow-up campaign for all clients.',
    ],
    insights: {
      quickWins: [
        'Automate consultation reminders in the CRM to reduce no-shows.',
        'Use the Fit2Go testimonial prompt kit to gather weekly social proof.',
      ],
      stretchMoves: [
        'Pitch a corporate wellness trial to an existing client’s employer.',
        'Co-create a client success story with marketing for the Fit2Go blog.',
      ],
    },
  },
  {
    name: 'Program Design & Periodization',
    description:
      'Translate client assessments into phased training plans that deliver measurable wins and sustainable habits.',
    trackerCaption: 'Update this tracker after each training block review.',
    tasks: [
      'Audit client baselines and map out primary adaptations targeted per block.',
      'Design a progression for strength, mobility, and recovery markers.',
      'Integrate behavior change practices that anchor habit adoption.',
    ],
    insights: {
      quickWins: [
        'Leverage the Fit2Go template library to accelerate programming.',
        'Pair each training phase with a 10-minute micro-education video.',
      ],
      stretchMoves: [
        'Collaborate with the nutrition team on a dual-coaching pilot.',
        'Create a client-friendly dashboard with before/after biometrics.',
      ],
    },
  },
  {
    name: 'Hybrid Coaching Experience',
    description:
      'Blend in-person energy with virtual accountability systems to drive unmatched client results.',
    trackerCaption: 'Mark your mastery each time you refine hybrid rituals.',
    tasks: [
      'Implement a “day-before” touchpoint workflow for hybrid clients.',
      'Refine your video coaching setup and lighting checklist.',
      'Test Fit2Go’s habit tracking beta with three clients.',
    ],
    insights: {
      quickWins: [
        'Adopt the hybrid client kickoff script to set expectations early.',
        'Batch record movement demos to reuse across clients.',
      ],
      stretchMoves: [
        'Lead a live virtual mobility lab for the Fit2Go community.',
        'Build a resource vault that blends workouts, recipes, and mindset tools.',
      ],
    },
  },
  {
    name: 'Leadership & Career Acceleration',
    description:
      'Level up as a Fit2Go leader by mentoring peers, influencing strategy, and shaping new offerings.',
    trackerCaption: 'Capture each leadership moment to build your promotion case.',
    tasks: [
      'Mentor a new trainer and log bi-weekly feedback loops.',
      'Facilitate a “Client Wins” roundtable for the regional team.',
      'Pitch a new initiative to Fit2Go leadership with forecasted outcomes.',
    ],
    insights: {
      quickWins: [
        'Share your top session prep system in the internal community.',
        'Volunteer for the next trainer recruitment day.',
      ],
      stretchMoves: [
        'Launch a specialty workshop track for high-performing clients.',
        'Co-lead a quarterly innovation sprint with operations.',
      ],
    },
  },
];

const seedMilestones = [
  {
    quarter: 'Q1 2024',
    title: 'Launch 90-day onboarding sprint',
    impact: 'Achieve 90%+ completion of onboarding tasks and secure first five client testimonials.',
  },
  {
    quarter: 'Q3 2024',
    title: 'Lead corporate wellness pilot',
    impact: 'Deliver hybrid workshops for Acme Tech, targeting 200 employee registrations.',
  },
  {
    quarter: 'Q1 2025',
    title: 'Mentor circle captain',
    impact: 'Support three new trainers and contribute to Fit2Go training playbook updates.',
  },
];

const journeyStartButton = document.getElementById('journeyStart');
const onboardingDetails = document.getElementById('onboardingDetails');
const onboardingTitle = document.getElementById('onboardingTitle');
const onboardingSummary = document.getElementById('onboardingSummary');
const steps = Array.from(document.querySelectorAll('.step'));
const stepTemplate = document.getElementById('stepTemplate');
const skillsGrid = document.getElementById('skillsGrid');
const skillCardTemplate = document.getElementById('skillCardTemplate');
const skillInsights = document.getElementById('skillInsights');
const timeline = document.getElementById('timeline');
const plannerForm = document.getElementById('plannerForm');

function renderOnboardingStep(index) {
  const step = onboardingSteps[index];
  onboardingTitle.textContent = step.title;
  onboardingSummary.textContent = step.summary;

  onboardingDetails.innerHTML = '';
  const checklist = stepTemplate.content.firstElementChild.cloneNode(true);

  step.tasks.forEach((task, taskIndex) => {
    const listItem = document.createElement('li');
    listItem.className = 'checklist__item';

    const checkboxId = `step-${index}-task-${taskIndex}`;
    listItem.innerHTML = `
      <input type="checkbox" id="${checkboxId}" />
      <label for="${checkboxId}">
        ${task.label}
        <span>${task.detail}</span>
      </label>
    `;
    checklist.appendChild(listItem);
  });

  onboardingDetails.appendChild(checklist);
}

function handleStepClick(event) {
  const button = event.currentTarget;
  const index = Number.parseInt(button.dataset.step, 10);
  steps.forEach((step) => step.classList.toggle('is-active', step === button));
  renderOnboardingStep(index);
}

steps.forEach((button) => {
  button.addEventListener('click', handleStepClick);
});

journeyStartButton?.addEventListener('click', () => {
  document.getElementById('onboarding').scrollIntoView({ behavior: 'smooth' });
});

function formatProgress(value, total) {
  return Math.round((value / total) * 100);
}

function renderSkillCard(area) {
  const node = skillCardTemplate.content.firstElementChild.cloneNode(true);
  const title = node.querySelector('.skill-card__title');
  const description = node.querySelector('.skill-card__description');
  const tasks = node.querySelector('.skill-card__tasks');
  const progress = node.querySelector('.skill-card__progress');
  const expandButton = node.querySelector('.skill-card__expand');
  const tracker = node.querySelector('.skill-card__tracker');
  const trackerCaption = node.querySelector('.skill-card__tracker-caption');
  const meter = node.querySelector('.skill-card__meter');
  const meterFill = node.querySelector('.skill-card__meter-fill');

  title.textContent = area.name;
  description.textContent = area.description;
  trackerCaption.textContent = area.trackerCaption;

  area.tasks.forEach((task) => {
    const li = document.createElement('li');
    li.textContent = task;
    tasks.appendChild(li);
  });

  const state = {
    completed: 0,
    total: area.tasks.length,
  };

  function updateProgressDisplay() {
    const pct = formatProgress(state.completed, state.total);
    progress.textContent = `${pct}% complete`;
    meter.setAttribute('aria-valuenow', String(pct));
    meterFill.style.width = `${pct}%`;
  }

  updateProgressDisplay();

  expandButton.addEventListener('click', (event) => {
    event.stopPropagation();
    const isExpanded = expandButton.getAttribute('aria-expanded') === 'true';
    expandButton.setAttribute('aria-expanded', String(!isExpanded));
    tracker.hidden = isExpanded;
    expandButton.textContent = isExpanded ? 'Mark progress' : 'Hide tracker';
  });

  node.addEventListener('click', () => {
    renderInsights(area);
  });

  node.addEventListener('keypress', (event) => {
    if (event.key === 'Enter' || event.key === ' ') {
      event.preventDefault();
      renderInsights(area);
    }
  });

  const completionChecklist = document.createElement('ul');
  completionChecklist.className = 'progress-list';
  completionChecklist.setAttribute('aria-label', `${area.name} completion tracker`);
  area.tasks.forEach((task, index) => {
    const item = document.createElement('li');
    const id = `${area.name.replace(/\s+/g, '-')}-progress-${index}`;
    item.innerHTML = `
      <label>
        <input type="checkbox" id="${id}" /> ${task}
      </label>
    `;
    const checkbox = item.querySelector('input');
    checkbox.addEventListener('change', () => {
      state.completed += checkbox.checked ? 1 : -1;
      updateProgressDisplay();
    });
    completionChecklist.appendChild(item);
  });

  tracker.appendChild(completionChecklist);

  return node;
}

function renderInsights(area) {
  skillInsights.innerHTML = `
    <h3>${area.name}</h3>
    <p>${area.description}</p>
    <h4>Quick wins</h4>
    <ul>
      ${area.insights.quickWins.map((item) => `<li>${item}</li>`).join('')}
    </ul>
    <h4>Stretch moves</h4>
    <ul>
      ${area.insights.stretchMoves.map((item) => `<li>${item}</li>`).join('')}
    </ul>
  `;
}

skillAreas.forEach((area, index) => {
  const card = renderSkillCard(area);
  if (index === 0) {
    renderInsights(area);
  }
  skillsGrid.appendChild(card);
});

function renderTimelineItem(milestone) {
  const item = document.createElement('article');
  item.className = 'timeline__item';
  item.innerHTML = `
    <span class="timeline__quarter">${milestone.quarter}</span>
    <h3 class="timeline__title">${milestone.title}</h3>
    <p class="timeline__impact">${milestone.impact}</p>
  `;
  return item;
}

function renderTimeline(milestones) {
  timeline.innerHTML = '';
  milestones.forEach((milestone) => {
    timeline.appendChild(renderTimelineItem(milestone));
  });
}

const milestones = [...seedMilestones];
renderTimeline(milestones);

plannerForm.addEventListener('submit', (event) => {
  event.preventDefault();
  const title = document.getElementById('milestoneTitle');
  const quarter = document.getElementById('milestoneQuarter');
  const impact = document.getElementById('milestoneImpact');

  milestones.push({
    title: title.value.trim(),
    quarter: quarter.value.trim(),
    impact: impact.value.trim() || 'Clarify impact areas during your next mentor sync.',
  });

  renderTimeline(milestones);
  plannerForm.reset();
  title.focus();
});

renderOnboardingStep(0);
renderInsights(skillAreas[0]);
