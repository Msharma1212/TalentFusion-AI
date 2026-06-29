import express from "express";
import path from "path";
import multer from "multer";
import dotenv from "dotenv";
import fs from "fs";
import os from "os";
import { execSync } from "child_process";
import { createServer as createViteServer } from "vite";
import { GoogleGenAI, Type } from "@google/genai";
import * as XLSX from "xlsx";
import { Candidate, TeamComposition, TeamCompatibility, SkillGapItem, InterviewQuestion } from "./src/types.js";

dotenv.config();

const app = express();
const PORT = 3000;

// Configure Multer for PDF resume uploads
const storage = multer.memoryStorage();
const upload = multer({
  storage,
  limits: { fileSize: 10 * 1024 * 1024 }, // 10MB limit
  fileFilter: (req, file, cb) => {
    if (file.mimetype === "application/pdf") {
      cb(null, true);
    } else {
      cb(new Error("Only PDF files are supported."));
    }
  }
});

// Initialize Gemini API Client
const geminiApiKey = process.env.GEMINI_API_KEY;
let aiClient: GoogleGenAI | null = null;

if (geminiApiKey) {
  aiClient = new GoogleGenAI({
    apiKey: geminiApiKey,
    httpOptions: {
      headers: {
        "User-Agent": "aistudio-build",
      }
    }
  });
}

// In-Memory Database of Candidates
let candidatesDb: Candidate[] = [
  {
    id: "cand-1",
    name: "Sarah Jenkins",
    skills: ["React", "TypeScript", "TailwindCSS", "Next.js", "Redux", "Figma", "Cypress", "Web Performance", "REST APIs"],
    experience: 6,
    projects: [
      {
        name: "Enterprise Design System",
        description: "Built a fully accessible, reusable UI component library used across 12 product lines, reducing frontend development time by 40%.",
        technologies: ["React", "TypeScript", "TailwindCSS", "Storybook"]
      },
      {
        name: "E-Commerce Replatforming",
        description: "Migrated a legacy multi-page site to a modern headless Next.js SPA, boosting mobile conversion rate by 25% and improving Lighthouse scores to 95+.",
        technologies: ["Next.js", "TypeScript", "GraphQL", "TailwindCSS"]
      }
    ],
    certifications: ["AWS Certified Cloud Practitioner", "Scrum Alliance Certified Developer"],
    education: "B.S. in Computer Science, University of Washington",
    leadership: 78,
    communication: 94,
    learningAbility: 88,
    summary: "Senior Frontend Engineer specializing in highly interactive, accessible, and performant web applications. Proven track record of bridging the gap between design and development by constructing polished component ecosystems.",
    recommendedRole: "Lead Frontend Engineer / UI Architect",
    strengths: ["Exemplary communication skills", "Deep expertise in modern React/Next.js ecosystem", "Strong design sense and UX advocacy"],
    weaknesses: ["Limited experience with server-side database tuning", "Prefers visual styling over pure data processing"],
    overallScore: 92,
    matchReason: "Exceptional frontend skill set, stellar design system leadership, and high communication score makes her a top candidate for leading user-facing product initiatives.",
    email: "s.jenkins@talentfusion.io",
    phone: "+1 (555) 382-9901",
    achievements: [
      "Designed and open-sourced 'SleekUI', a Tailwind CSS component library with over 15k GitHub stars.",
      "Spearheaded frontend migration of Next.js checkout system, increasing checkout completion by 24%."
    ]
  },
  {
    id: "cand-2",
    name: "David Chen",
    skills: ["Go", "Node.js", "PostgreSQL", "Redis", "Docker", "Kubernetes", "AWS", "gRPC", "Microservices", "System Design"],
    experience: 8,
    projects: [
      {
        name: "High-Throughput Payment Gateway",
        description: "Architected a transaction processing backend in Go capable of handling 15,000 requests per second with 99.99% uptime.",
        technologies: ["Go", "PostgreSQL", "Redis", "Kafka", "Docker"]
      },
      {
        name: "Legacy Monolith Migration",
        description: "Successfully decoupled a giant Express monolith into 14 Dockerized gRPC microservices orchestrating via Kubernetes.",
        technologies: ["Node.js", "Go", "Kubernetes", "AWS", "gRPC"]
      }
    ],
    certifications: ["Certified Kubernetes Administrator (CKA)", "AWS Solutions Architect Professional"],
    education: "B.S. in Software Engineering, UT Austin",
    leadership: 85,
    communication: 72,
    learningAbility: 90,
    summary: "Senior Backend Architect dedicated to constructing highly resilient, scalable server architectures, database modeling, and microservice infrastructure.",
    recommendedRole: "Principal Backend Engineer / System Architect",
    strengths: ["Master of concurrent system designs", "Extensive experience with containers and orchestration", "Strong database query optimization"],
    weaknesses: ["Communication can sometimes be overly clinical", "Minimal UI/UX design or styling experience"],
    overallScore: 95,
    matchReason: "A technical powerhouse who excels at data pipelines, microservices, and system reliability. Complements team members with weaker systems/operations knowledge.",
    email: "d.chen@talentfusion.io",
    phone: "+1 (555) 472-8311",
    achievements: [
      "Optimized core payment gateway throughput, reducing API latency from 240ms to 12ms under full load.",
      "Keynote speaker at GopherCon 2024 on high-frequency transaction logging in Go."
    ]
  },
  {
    id: "cand-3",
    name: "Dr. Elena Rostova",
    skills: ["Python", "PyTorch", "TensorFlow", "Transformers", "LLMs", "RAG", "Data Pipelines", "Machine Learning", "FastAPI", "SQL"],
    experience: 5,
    projects: [
      {
        name: "Corporate Intelligence Copilot",
        description: "Developed an agentic RAG system that indexes over 500,000 corporate documents, enabling context-aware semantic search and summary generation.",
        technologies: ["Python", "PyTorch", "LangChain", "VectorDB", "FastAPI"]
      },
      {
        name: "Predictive Customer Churn Model",
        description: "Engineered and deployed a pipeline of gradient-boosted trees predicting customer attrition with 89% precision, saving $2M in annual revenue.",
        technologies: ["Python", "Scikit-Learn", "XGBoost", "MLflow"]
      }
    ],
    certifications: ["Google Cloud Professional Machine Learning Engineer", "DeepLearning.AI Generative AI Specialist"],
    education: "Ph.D. in Computer Science (AI Focus), Stanford University",
    leadership: 65,
    communication: 82,
    learningAbility: 98,
    summary: "AI Scientist with a Ph.D. specializing in large language model operations, retrieval-augmented generation (RAG), and deploying production-ready machine learning pipelines.",
    recommendedRole: "Lead AI/ML Engineer",
    strengths: ["World-class cognitive ability and mathematical foundations", "Stellar knowledge of generative AI models and custom embeddings", "Quick adapter of modern research methodologies"],
    weaknesses: ["Can sometimes deep-dive into research over rapid product delivery", "Prefers command-line tools over user interfaces"],
    overallScore: 96,
    matchReason: "Exceptional theoretical and practical machine learning credentials. Crucial hire for integrating advanced cognitive reasoning into software solutions.",
    email: "e.rostova@talentfusion.io",
    phone: "+1 (555) 902-7721",
    achievements: [
      "Authored 4 peer-reviewed ACL publications on multi-modal Retrieval-Augmented Generation.",
      "Built and tuned customized 14B parameter llama-finetuned model with 92% retrieval precision on proprietary law corpora."
    ]
  },
  {
    id: "cand-4",
    name: "Marcus Sterling",
    skills: ["Figma", "User Research", "Wireframing", "Prototyping", "Information Architecture", "Design Systems", "HTML/CSS", "Adobe CC", "TailwindCSS"],
    experience: 4,
    projects: [
      {
        name: "Mobile Wallet Experience Re-design",
        description: "Led the complete user research and end-to-end design of a fintech mobile app, reducing user onboarding friction by 60%.",
        technologies: ["Figma", "Miro", "UserTesting"]
      },
      {
        name: "Design Token Architecture",
        description: "Established automated design token pipelines bridging Figma styles directly to Tailwind/JSON components, ensuring flawless brand alignment.",
        technologies: ["Figma", "Style Dictionary", "JSON", "GitHub"]
      }
    ],
    certifications: ["Google UX Design Professional Certificate", "NN/g UX Certified"],
    education: "B.A. in Interaction Design, Georgia Tech",
    leadership: 70,
    communication: 96,
    learningAbility: 85,
    summary: "Product Designer passionate about user centered methodologies, interactive prototypes, and constructing scalable, pixel-perfect design languages.",
    recommendedRole: "Lead Product/UI Designer",
    strengths: ["Exceptional empathy and user advocacy", "Stellar presentation and client collaboration skills", "Seamless collaboration with engineering workflows"],
    weaknesses: ["Does not write production backend logic", "Focuses heavily on pixel perfection which can delay early-stage wireframes"],
    overallScore: 89,
    matchReason: "Provides crucial user advocacy, rapid wireframing capability, and absolute brand coherence. Blends wonderfully with high-velocity engineers.",
    email: "m.sterling@talentfusion.io",
    phone: "+1 (555) 233-1082",
    achievements: [
      "Redesigned main mobile banking portal, leading to a 4.8 star App Store rating (up from 3.2).",
      "Pioneered style-dictionary automation matching Figma designer nodes directly to CSS custom properties."
    ]
  },
  {
    id: "cand-5",
    name: "Rajesh Kumar",
    skills: ["React", "Node.js", "Express", "TypeScript", "PostgreSQL", "Docker", "GCP", "CI/CD", "System Architecture", "Project Management"],
    experience: 7,
    projects: [
      {
        name: "SaaS Project Management Hub",
        description: "Acted as technical lead for a team of 5 engineers delivering a real-time collaborative workspace, utilizing WebSockets for real-time sync.",
        technologies: ["React", "Node.js", "Socket.io", "PostgreSQL", "GCP"]
      },
      {
        name: "Automated Deploy Pipeline",
        description: "Configured multi-stage CI/CD pipelines reducing feature deployment cycle times from 3 days to under 15 minutes with automated unit test suites.",
        technologies: ["GitHub Actions", "Docker", "GCP Cloud Run", "Bash"]
      }
    ],
    certifications: ["Google Cloud Certified Professional Cloud Architect", "Certified ScrumMaster (CSM)"],
    education: "B.S. in Information Technology, IIT Kharagpur",
    leadership: 92,
    communication: 90,
    learningAbility: 86,
    summary: "Technical Leader and Full Stack Developer with 7 years of engineering experience. Combines deep system understanding with modern Scrum leadership practices.",
    recommendedRole: "Technical Team Lead / Engineering Manager",
    strengths: ["Superb Scrum leadership and organizational clarity", "Versatile Full Stack coverage", "Excellent mentor for junior developers"],
    weaknesses: ["Spreads focus across management and coding, leaving deep niche tuning to specialists", "Can be cautious regarding high-risk experimental architectures"],
    overallScore: 94,
    matchReason: "The perfect candidate to bridge executive goals with daily engineering velocity. His high leadership and Full Stack capabilities bind the team together.",
    email: "r.kumar@talentfusion.io",
    phone: "+1 (555) 762-0941",
    achievements: [
      "Delivered multi-tenant collaboration suite 3 weeks ahead of schedule, with zero critical defects.",
      "Successfully scaled engineering operations from 3 to 14 cross-functional squad members."
    ]
  }
];

// Base list of original candidates to restore on Reset
const ORIGINAL_CANDIDATES = [...candidatesDb];

// ==========================================
// ROBUST LOCAL BACKUP/FALLBACK GENERATORS
// ==========================================

function parsePdfFallback(fileName: string, jobDescription: string): Candidate {
  const cleanName = fileName
    .replace(/\.[^/.]+$/, "") // remove extension
    .replace(/[-_]/g, " ")    // replace dashes/underscores with space
    .replace(/\b\w/g, c => c.toUpperCase()); // Capitalize first letters of words

  const jdLower = (jobDescription || "").toLowerCase();

  let matchedRole = "Senior Full Stack Engineer";
  let skills = ["React", "TypeScript", "Node.js", "Express", "PostgreSQL", "TailwindCSS", "REST APIs", "Docker", "Git"];
  let summary = `Seasoned engineer with a focus on core TypeScript architectures, backend integrations, and responsive frontend dashboards. Proven record of deploying scalable services.`;

  if (jdLower.includes("ai") || jdLower.includes("ml") || jdLower.includes("python") || jdLower.includes("model")) {
    matchedRole = "AI/ML Solutions Engineer";
    skills = ["Python", "PyTorch", "Hugging Face", "FastAPI", "Docker", "SQL", "Transformers", "LLMs", "RAG"];
    summary = `AI Engineer specialized in integrating large language models, developing vector database indexing pipelines (RAG), and serving robust web-accessible inference layers.`;
  } else if (jdLower.includes("design") || jdLower.includes("ui") || jdLower.includes("ux") || jdLower.includes("figma") || jdLower.includes("frontend")) {
    matchedRole = "Senior UI & Design Systems Engineer";
    skills = ["React", "TypeScript", "Figma", "TailwindCSS", "HTML/CSS", "Next.js", "Jest", "Storybook", "User Experience"];
    summary = `Polished frontend developer dedicated to bridging design wireframes with clean, accessible, and performant user interface component libraries.`;
  }

  return {
    id: `cand-fallback-${Date.now()}`,
    name: cleanName,
    skills,
    experience: Math.floor(Math.random() * 5) + 4, // 4-8 yrs
    projects: [
      {
        name: "Enterprise Scalability Initiative",
        description: "Orchestrated performance optimizations on core workflows, boosting delivery speeds and decoupling bottleneck dependencies.",
        technologies: skills.slice(0, 4)
      },
      {
        name: "Custom Integration Dashboard",
        description: "Architected a responsive interactive portal enabling seamless search, visualization, and monitoring of platform datasets.",
        technologies: skills.slice(2, 6)
      }
    ],
    certifications: ["Professional Scrum Developer", "Cloud Native Certified Associate"],
    education: "B.S. in Computer Science",
    leadership: Math.floor(Math.random() * 20) + 70, // 70-90
    communication: Math.floor(Math.random() * 20) + 75, // 75-95
    learningAbility: Math.floor(Math.random() * 15) + 80, // 80-95
    summary,
    recommendedRole: matchedRole,
    strengths: ["Excellent team alignment and technical execution", "Pragmatic balance of speed and codebase health", "Strong user empathy and structured thinking"],
    weaknesses: ["Can sometimes be conservative with highly experimental alpha releases", "Aims for perfect unit test coverage which can delay immediate feedback"],
    overallScore: Math.floor(Math.random() * 10) + 85, // 85-95
    matchReason: `Generated via local parsing engine fallback. Profile indicates a superb core match for the job description, with outstanding alignment in TypeScript-based web engineering, reliable full-stack delivery protocols, and proactive communication habits.`
  };
}

function buildSingleTeamFallback(
  candidates: Candidate[],
  requirements: string,
  teamSize: number,
  archetype: "best" | "alternative" | "budget" | "innovation" | "delivery"
): any {
  const reqLower = (requirements || "").toLowerCase();
  let sorted = [...candidates];

  if (archetype === "best") {
    sorted.sort((a, b) => b.overallScore - a.overallScore);
  } else if (archetype === "alternative") {
    // Exclude the absolute best overall candidate to force an alternative talent structure
    sorted.sort((a, b) => b.overallScore - a.overallScore);
    if (sorted.length > 2) {
      sorted.splice(0, 1);
    }
  } else if (archetype === "budget") {
    // Prefer highly skilled candidates with moderate experience (cost-efficiency)
    sorted.sort((a, b) => {
      const valB = b.overallScore / (b.experience + 1);
      const valA = a.overallScore / (a.experience + 1);
      return valB - valA;
    });
  } else if (archetype === "innovation") {
    // Bias towards high learning ability and specialized AI/ML or Design skills
    sorted.sort((a, b) => {
      const isAiOrDesign = (c: Candidate) => c.skills.some(s => 
        ["python", "figma", "pytorch", "transformers", "rag", "llms", "user research"].includes(s.toLowerCase())
      );
      const scoreA = a.learningAbility * 1.5 + (isAiOrDesign(a) ? 40 : 0);
      const scoreB = b.learningAbility * 1.5 + (isAiOrDesign(b) ? 40 : 0);
      return scoreB - scoreA;
    });
  } else if (archetype === "delivery") {
    // Bias towards solid experience, leadership, and system engineering architectures
    sorted.sort((a, b) => {
      const scoreA = a.experience * 12 + a.leadership * 1.2;
      const scoreB = b.experience * 12 + b.leadership * 1.2;
      return scoreB - scoreA;
    });
  }

  const size = Math.min(teamSize, sorted.length);
  const selectedCandidates = sorted.slice(0, size);

  if (selectedCandidates.length === 0) {
    return {
      id: `team-composition-empty-${Date.now()}`,
      name: "Empty Team",
      roleAssignments: [],
      compatibility: {
        skillDiversity: 0,
        leadershipBalance: 0,
        experienceBalance: 0,
        communication: 0,
        learningPotential: 0,
        conflictRisk: 0,
        innovationScore: 0,
        deliveryConfidence: 0,
        overall: 0,
        explanation: "No candidates available.",
        burnoutRisk: 0,
        knowledgeCoverage: 0,
        communicationBalance: 0
      },
      strengths: [],
      risks: []
    };
  }

  // Generate role assignments
  const roleAssignments = selectedCandidates.map(c => {
    let assignedRole = c.recommendedRole;
    if (reqLower.includes("backend") && c.skills.some(s => s.toLowerCase() === "go")) {
      assignedRole = "Cloud Systems Developer";
    } else if (reqLower.includes("frontend") && c.skills.some(s => s.toLowerCase() === "react")) {
      assignedRole = "User Interface Architect";
    } else if (reqLower.includes("ai") && c.skills.some(s => s.toLowerCase() === "python")) {
      assignedRole = "AI/ML Solutions Lead";
    } else if (reqLower.includes("lead") && c.leadership > 80) {
      assignedRole = "Technical Program Director";
    } else if (reqLower.includes("designer") && c.skills.some(s => s.toLowerCase() === "figma")) {
      assignedRole = "Experience Design Lead";
    }
    return {
      role: assignedRole,
      candidateId: c.id,
      candidateName: c.name
    };
  });

  const averageComm = Math.round(selectedCandidates.reduce((acc, c) => acc + c.communication, 0) / size);
  const averageLead = Math.round(selectedCandidates.reduce((acc, c) => acc + c.leadership, 0) / size);
  const averageLearn = Math.round(selectedCandidates.reduce((acc, c) => acc + c.learningAbility, 0) / size);
  const averageExp = selectedCandidates.reduce((acc, c) => acc + c.experience, 0) / size;

  const hasAI = selectedCandidates.some(c => c.skills.some(s => ["python", "pytorch", "transformers", "llms", "rag"].includes(s.toLowerCase())));
  const hasUI = selectedCandidates.some(c => c.skills.some(s => ["react", "figma", "tailwindcss", "next.js"].includes(s.toLowerCase())));
  const hasBackend = selectedCandidates.some(c => c.skills.some(s => ["go", "node.js", "postgresql", "redis"].includes(s.toLowerCase())));
  const hasLeader = selectedCandidates.some(c => c.leadership >= 85);

  const skillDiversity = Math.min(98, 70 + (Array.from(new Set(selectedCandidates.flatMap(c => c.skills))).length * 1.5));
  const leadershipBalance = hasLeader ? (selectedCandidates.filter(c => c.leadership >= 85).length === 1 ? 95 : 82) : 60;
  const experienceBalance = Math.round(100 - Math.abs(averageExp - 6.0) * 5);
  const communication = averageComm;
  const learningPotential = averageLearn;
  const conflictRisk = Math.max(12, Math.round(100 - (averageComm * 1.15 - (averageLead - 75) * 0.15)));
  const innovationScore = Math.min(98, 55 + (hasAI ? 25 : 0) + (hasUI ? 15 : 0) + (averageLearn >= 90 ? 10 : 0));
  const deliveryConfidence = Math.min(98, 50 + (hasBackend ? 20 : 0) + (hasLeader ? 15 : 0) + (averageExp >= 6 ? 15 : 0));

  // chemistry fields
  const burnoutRisk = Math.round(30 + (averageLead * 0.25) - (averageExp * 2.0) + (Math.random() * 8));
  const knowledgeCoverage = Math.min(98, 50 + (Array.from(new Set(selectedCandidates.flatMap(c => c.skills))).length * 2));
  const communicationBalance = Math.round(100 - Math.abs(averageComm - 85) * 1.5);

  const overall = Math.round((skillDiversity + leadershipBalance + experienceBalance + communication + learningPotential + (100 - conflictRisk) + innovationScore + deliveryConfidence) / 8);

  // Set names based on archetype
  let name = "";
  if (archetype === "best") {
    name = "Team Apex Synergy";
  } else if (archetype === "alternative") {
    name = "Team Paradigm Shift";
  } else if (archetype === "budget") {
    name = "Cost-Optimized Taskforce";
  } else if (archetype === "innovation") {
    name = "Synapse AI Lab";
  } else if (archetype === "delivery") {
    name = "Velocity Delivery Command";
  }

  // Set explanations
  let explanation = "";
  if (archetype === "best") {
    explanation = `The supreme technical selection. Features pristine balance of product architecture, pixel-perfect layouts, and advanced machine learning modeling, backed by an average of ${averageExp.toFixed(1)} years of experience.`;
  } else if (archetype === "alternative") {
    explanation = `A premium backup crew omitting the highest scoring candidates. Highly aligned on responsive communication and rapid iterations, securing project continuity.`;
  } else if (archetype === "budget") {
    explanation = `Maximized ROI cluster. Combines highly active modern specialists with high cognitive learning speed, delivering robust features without senior-heavy payroll overheads.`;
  } else if (archetype === "innovation") {
    explanation = `Engineered for bleeding-edge experiments. Leverages world-class Stanford Ph.D. capabilities and style-token automations, scoring an outstanding ${innovationScore}% on creative R&D.`;
  } else if (archetype === "delivery") {
    explanation = `Built for high-throughput, mission-critical product launches. Heavy system engineering bias with senior-level orchestration guarantees 99.99% operational uptimes.`;
  }

  // DNA report fields
  const recommendedLeadCand = selectedCandidates.reduce((max, c) => c.leadership > max.leadership ? c : max, selectedCandidates[0]);
  const workingStyle = archetype === "innovation" ? "R&D Prototyping Sprints" : (archetype === "delivery" ? "Continuous Integration Taskforce" : "Hybrid Agile SCRUM");
  
  const teamDna = {
    innovationIndex: innovationScore,
    executionIndex: deliveryConfidence,
    leadershipIndex: averageLead,
    learningSpeed: averageLearn,
    communicationIndex: averageComm,
    technicalCoverage: skillDiversity,
    riskLevel: (conflictRisk > 35 ? "Medium" : "Low") as "Low" | "Medium" | "High",
    recommendedLead: recommendedLeadCand.name,
    workingStyle
  };

  const strengths = [];
  if (hasAI && hasUI) strengths.push("Unlocks state-of-the-art cognitive interfaces and AI application cycles.");
  if (hasBackend && hasUI) strengths.push("Flawless full-stack delivery capabilities from datastore to pixel.");
  if (averageComm >= 85) strengths.push("Exceptional structural cohesion with stellar internal communication rates.");
  if (strengths.length < 3) strengths.push("Highly adaptable skill sets capable of picking up complex workflows rapidly.");

  const risks = [];
  if (!hasAI) risks.push("Limited generative cognitive experience, might require specialized model consultants.");
  if (!hasUI) risks.push("Heavy operational backend bias which could lead to secondary UX adoption friction.");
  if (!hasBackend) risks.push("Absence of specialized high-throughput systems infrastructure engineers.");
  if (risks.length < 2) risks.push("Accelerated onboarding learning curve during the initial two-week sprint phase.");

  return {
    id: `team-composition-${archetype}-${Date.now()}`,
    name,
    roleAssignments,
    compatibility: {
      skillDiversity,
      leadershipBalance,
      experienceBalance,
      communication,
      learningPotential,
      conflictRisk,
      innovationScore,
      deliveryConfidence,
      overall,
      explanation,
      burnoutRisk,
      knowledgeCoverage,
      communicationBalance
    },
    strengths,
    risks,
    teamDna
  };
}

function buildTeamFallback(candidates: Candidate[], requirements: string, teamSize: number): any {
  return {
    best: buildSingleTeamFallback(candidates, requirements, teamSize, "best"),
    alternative: buildSingleTeamFallback(candidates, requirements, teamSize, "alternative"),
    budget: buildSingleTeamFallback(candidates, requirements, teamSize, "budget"),
    innovation: buildSingleTeamFallback(candidates, requirements, teamSize, "innovation"),
    delivery: buildSingleTeamFallback(candidates, requirements, teamSize, "delivery")
  };
}

function analyzeCandidateFallback(candidate: Candidate, targetRole: string, jobDescription: string): any {
  const role = targetRole || candidate.recommendedRole;
  const jdLower = (jobDescription || "").toLowerCase();

  // Calculate score variation based on job description alignment
  let scoreModifier = 0;
  candidate.skills.forEach(s => {
    if (jdLower.includes(s.toLowerCase())) {
      scoreModifier += 2;
    }
  });
  const finalScore = Math.min(100, Math.max(50, candidate.overallScore + scoreModifier));

  // Dynamic Skill Gaps
  const skillGaps: SkillGapItem[] = [];
  const candidateSkillsLower = candidate.skills.map(s => s.toLowerCase());

  // Check common software/designer/AI gaps
  if (candidateSkillsLower.includes("react")) {
    // Frontend developer gaps
    skillGaps.push({
      skill: "High-Throughput Backend Architecture",
      status: "missing",
      roadmap: [
        {
          resource: "Go Programming Language Tour & Designing Data-Intensive Applications",
          estimatedDays: 14,
          description: "Establish familiarity with strong concurrent typing, system design boundaries, and relational schemas."
        },
        {
          resource: "Pragmatic Microservices in Practice (O'Reilly)",
          estimatedDays: 10,
          description: "Build an asynchronous message-passing queue using Go, Docker, and PostgreSQL."
        }
      ]
    });
    skillGaps.push({
      skill: "Advanced Cloud & Kubernetes Orchestration",
      status: "partial",
      roadmap: [
        {
          resource: "AWS Certified Solutions Architect Course",
          estimatedDays: 12,
          description: "Gain hands-on knowledge of cloud scaling, ECS, security boundaries, and serverless compute."
        }
      ]
    });
  } else if (candidateSkillsLower.includes("go") || candidateSkillsLower.includes("node.js")) {
    // Backend gaps
    skillGaps.push({
      skill: "Modern UI Component Design Systems",
      status: "missing",
      roadmap: [
        {
          resource: "Tailwind CSS Documentation & Figma Interactive Guides",
          estimatedDays: 7,
          description: "Master modern utility-first responsive layout paradigms, typography guidelines, and design systems token syncing."
        },
        {
          resource: "React - Complete Developer Guide (Academind)",
          estimatedDays: 14,
          description: "Construct interactive single-page dashboards utilizing atomic design structures, standard React state hooks, and client-side optimization."
        }
      ]
    });
  } else if (candidateSkillsLower.includes("python") || candidateSkillsLower.includes("pytorch")) {
    // AI Specialist gaps
    skillGaps.push({
      skill: "Enterprise System Deployment & DevOps",
      status: "partial",
      roadmap: [
        {
          resource: "Docker and Kubernetes Fundamentals (Linux Foundation)",
          estimatedDays: 8,
          description: "Containerize Python machine learning inference servers and setup automated Kubernetes horizontal pod autoscaling."
        },
        {
          resource: "Full-Stack Web Development for AI Solutions",
          estimatedDays: 10,
          description: "Connect FastAPIs seamlessly to modern Vite + React interfaces using streaming server-sent events for LLM text generation."
        }
      ]
    });
  } else {
    // Default Designer or generic gaps
    skillGaps.push({
      skill: "Production Code Integration",
      status: "partial",
      roadmap: [
        {
          resource: "HTML5, CSS3, & Tailwind UI Guidelines",
          estimatedDays: 10,
          description: "Transition wireframes into fully semantic HTML/CSS code directly mapped to Figma styles."
        }
      ]
    });
  }

  // Interview Questions tailored to candidate's background
  const interviewQuestions: InterviewQuestion[] = [];
  if (candidateSkillsLower.includes("react")) {
    interviewQuestions.push({
      type: "technical",
      question: "How do you manage complex rendering performance issues in extremely large React applications with nested component trees?",
      expectedAnswer: "Explain the usage of profiling, React DevTools, React.memo for component caching, useCallback/useMemo to prevent reference recreation, and proper state relocation or windowing (using virtual lists) for rendering large datasets.",
      difficulty: "Hard"
    });
    interviewQuestions.push({
      type: "technical",
      question: "What is your workflow when coordinating with backend engineers to integrate REST/GraphQL payloads securely?",
      expectedAnswer: "Advocate for early API specification contracts (using Swagger or OpenAPI), setting up mock servers immediately, establishing comprehensive TypeScript type mappings, and implementing strict client-side validation.",
      difficulty: "Medium"
    });
  } else {
    interviewQuestions.push({
      type: "technical",
      question: "How do you optimize slow SQL queries in a production database experiencing high transaction volumes?",
      expectedAnswer: "Explain EXPLAIN ANALYZE command usage, proper indexing structures, query profiling, denormalization, database connection pooling parameters, and setting up secondary read replicas for heavy queries.",
      difficulty: "Hard"
    });
    interviewQuestions.push({
      type: "technical",
      question: "What strategies do you adopt to secure microservice boundaries against unauthorized calls?",
      expectedAnswer: "Discuss mutual TLS (mTLS) inside the mesh, JSON Web Tokens (JWT) verified at the gateway level, secure environment variables handling, and role-based access control policies.",
      difficulty: "Medium"
    });
  }

  interviewQuestions.push({
    type: "behavioral",
    question: `Your profile notes a communication score of ${candidate.communication}/100. Tell us about a time you had to resolve a high-stakes technical disagreement on your team.`,
    expectedAnswer: "Look for strong active listening skills, presenting objective data and performance comparisons (such as Lighthouse stats or memory benchmarks), and seeking a compromise rather than imposing ideas.",
    difficulty: "Medium"
  });
  interviewQuestions.push({
    type: "behavioral",
    question: "Give an example of a goal you set but failed to reach. How did you pivot, and what did you learn?",
    expectedAnswer: "Candidates should demonstrate strong self-awareness, taking extreme ownership of the failure, communicating early with key stakeholders, and incorporating lessons into their next project cycle.",
    difficulty: "Easy"
  });

  interviewQuestions.push({
    type: "situational",
    question: "A critical feature launch is 48 hours away, and the QA team identifies a regression that impacts 5% of checkout users. Do you patch it immediately or delay the launch?",
    expectedAnswer: "Assess risk tolerance: weigh the business impact of a 5% failure against the delay. Standard best practice is rolling back the specific feature branch, or implementing a feature flag toggle, rather than rushing a last-minute hotfix without full regression test coverage.",
    difficulty: "Hard"
  });
  interviewQuestions.push({
    type: "situational",
    question: "The product manager asks to implement a feature that you strongly believe degrades the core user experience. How do you approach this conversation?",
    expectedAnswer: "Focus on user-centered objective feedback: provide mockups or screen flows demonstrating the user friction, suggest an alternative design that meets the same business objective with cleaner UX, and suggest running an A/B test.",
    difficulty: "Medium"
  });

  return {
    candidateId: candidate.id,
    candidateName: candidate.name,
    score: finalScore,
    roleAlignment: `Excellent alignment as a ${role}. Match score of ${finalScore}% is driven by robust real-world experience (${candidate.experience} years) and exceptional competency across their core tools.`,
    strengths: candidate.strengths,
    weaknesses: candidate.weaknesses,
    skillGaps,
    interviewQuestions
  };
}

function chatFallback(messages: any[], query: string): any {
  const text = query || (messages && messages[messages.length - 1]?.text) || "";
  const queryLower = text.toLowerCase();

  let replyText = "";
  let suggestedAction: any = undefined;
  let candidateIds: string[] = [];

  // Find matching candidates based on simple text search to make fallback smart
  const queryWords = queryLower.split(/[\s,]+/).filter(w => w.length > 2);
  const matchedCandidateIds = candidatesDb.filter(c => {
    return queryWords.some(w => 
      c.name.toLowerCase().includes(w) ||
      c.skills.some(s => s.toLowerCase().includes(w)) ||
      c.recommendedRole.toLowerCase().includes(w)
    );
  }).map(c => c.id);

  // Let's analyze common recruitment questions
  if (queryLower.includes("react") || queryLower.includes("frontend") || queryLower.includes("next.js") || queryLower.includes("design system")) {
    replyText = `Based on your request, I scanned our active registry and identified <b>Sarah Jenkins</b> as our premier Frontend & Design System expert. She has 6 years of experience, a 92% JD match rating, and holds deep expertise in React, TypeScript, Next.js, and Figma.<br/><br/>I have loaded the matching candidates below. Click any card to view their profile, compare them, or add them to the team builder sandbox.`;
    suggestedAction = {
      type: "view_candidate",
      payload: { id: "cand-1" }
    };
    candidateIds = ["cand-1", "cand-5"];
  } else if (queryLower.includes("go") || queryLower.includes("backend") || queryLower.includes("postgres") || queryLower.includes("database") || queryLower.includes("docker")) {
    replyText = `I found <b>David Chen</b>, who is our expert systems and database architect (8 years of Go, PostgreSQL, Redis, Kubernetes, AWS experience; 95% Match Score). Rajesh Kumar also has strong full-stack capability.<br/><br/>I have loaded the matching candidates below. Click any card to view their profile, compare them, or add them to the team builder sandbox.`;
    suggestedAction = {
      type: "view_candidate",
      payload: { id: "cand-2" }
    };
    candidateIds = ["cand-2", "cand-5"];
  } else if (queryLower.includes("ai") || queryLower.includes("ml") || queryLower.includes("python") || queryLower.includes("pytorch") || queryLower.includes("rag") || queryLower.includes("llm")) {
    replyText = `Our highest matching specialist is <b>Dr. Elena Rostova</b>, an expert AI Scientist with a Ph.D. from Stanford. She has extensive experience deploying production-ready machine learning pipelines, LLMs, and agentic RAG architectures (96% Match Score). <br/><br/>Review the matching candidate cards below to compare suitability, view gap analysis, or drag them directly into the Team Builder.`;
    suggestedAction = {
      type: "view_candidate",
      payload: { id: "cand-3" }
    };
    candidateIds = ["cand-3", "cand-4"];
  } else if (queryLower.includes("lead") || queryLower.includes("manager") || queryLower.includes("leaders")) {
    replyText = `I recommend <b>Rajesh Kumar</b>, our Technical Team Lead and Engineering Manager (7 years experience, 92% Leadership score, Certified ScrumMaster). Sarah Jenkins is also well-suited for a Lead Frontend role.<br/><br/>Review the matching candidate cards below to compare suitability, view gap analysis, or drag them directly into the Team Builder.`;
    suggestedAction = {
      type: "view_candidate",
      payload: { id: "cand-5" }
    };
    candidateIds = ["cand-5", "cand-1", "cand-2"];
  } else if (queryLower.includes("team") || queryLower.includes("build") || queryLower.includes("launch") || queryLower.includes("composition") || queryLower.includes("builder")) {
    replyText = `I am initiating the <b>AI Team Builder engine</b> to assemble an optimal cross-functional product development team. I will redirect you to the Team Builder workspace where you can review compatibility metrics, leadership balances, strengths, and risks.`;
    suggestedAction = {
      type: "build_team",
      payload: { requirements: text }
    };
    candidateIds = [];
  } else if (queryLower.includes("graph") || queryLower.includes("map") || queryLower.includes("network")) {
    replyText = `Switching to the <b>Talent Graph Map</b>. This is a dynamic visual network detailing candidate experience levels, overlap in technical skill pools, and relative alignment rates. Click on any node to select a candidate.`;
    suggestedAction = {
      type: "view_graph",
      payload: {}
    };
    candidateIds = [];
  } else {
    // General chat matching
    if (matchedCandidateIds.length > 0) {
      replyText = `Based on your request, I identified <b>${matchedCandidateIds.length} candidate(s)</b> matching your search criteria: <b>"${text}"</b>. Below are their interactive compatibility cards, containing fit scores, skill overlaps, and detailed assessments.`;
      candidateIds = matchedCandidateIds;
    } else {
      replyText = `Hello! I am <b>FusionAI</b>, your interactive workforce coordinator. I can help you search talent pools or construct optimal cross-functional teams.<br/><br/>Here are some quick actions you can try:<br/>• "Find candidates with React, Go, and AWS skills." (View Frontend or Backend specialists)<br/>• "Show candidates suitable for Team Lead." (Identify management talent)<br/>• "Build the best AI product launch team." (Trigger Team Builder)<br/>• "Show me the Talent Graph map" (Activate network overview)`;
      candidateIds = [];
    }
  }

  return {
    text: replyText,
    suggestedAction,
    candidateIds
  };
}

app.use(express.json());

// API: Health Check
app.get("/api/health", (req, res) => {
  res.json({ status: "ok", time: new Date() });
});

// API: List Candidates
app.get("/api/candidates", (req, res) => {
  res.json({ success: true, data: candidatesDb });
});

// API: Reset DB to original mock data
app.post("/api/candidates/reset", (req, res) => {
  candidatesDb = JSON.parse(JSON.stringify(ORIGINAL_CANDIDATES));
  res.json({ success: true, message: "Database reset to original mock candidates.", data: candidatesDb });
});

// API: Parse PDF Resume using Gemini Multimodal native input
app.post("/api/upload", upload.single("resume"), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ success: false, error: "No file uploaded." });
    }

    const jobDescription = req.body.jobDescription || "Full Stack Software Developer team member.";
    const originalName = req.file.originalname;

    if (!aiClient) {
      return res.status(500).json({ success: false, error: "Gemini API client not initialized. Cannot parse resume." });
    }

    // 1. Write the PDF buffer to a unique temporary file
    const tempPath = path.join(os.tmpdir(), `resume-${Date.now()}-${Math.floor(Math.random() * 1000)}.pdf`);
    fs.writeFileSync(tempPath, req.file.buffer);

    let extractedText = "";
    let parseMethod = "";

    // 2. Extract text using PyMuPDF and pdfplumber in Python (Rule 2)
    try {
      const output = execSync(`python3 /parser.py "${tempPath}"`, { encoding: "utf8" });
      if (output.includes("---METHOD:PYMUPDF---")) {
        parseMethod = "PyMuPDF";
        extractedText = output.split("---METHOD:PYMUPDF---")[1].trim();
      } else if (output.includes("---METHOD:PDFPLUMBER---")) {
        parseMethod = "pdfplumber";
        extractedText = output.split("---METHOD:PDFPLUMBER---")[1].trim();
      }
    } catch (pyErr: any) {
      console.warn("Python-based text extraction failed:", pyErr.message || pyErr);
    } finally {
      try {
        fs.unlinkSync(tempPath);
      } catch (unlinkErr) {
        console.error("Failed to remove temp file:", unlinkErr);
      }
    }

    // 3. Fallback to OCR if Python extraction is blank/too short (Rule 2)
    if (!extractedText || extractedText.length < 50) {
      try {
        console.log("Python PDF text extraction returned empty or failed. Invoking Gemini OCR fallback...");
        parseMethod = "Gemini OCR";
        const pdfBase64 = req.file.buffer.toString("base64");
        const ocrResponse = await aiClient.models.generateContent({
          model: "gemini-3.5-flash",
          contents: [
            {
              inlineData: {
                data: pdfBase64,
                mimeType: "application/pdf"
              }
            },
            {
              text: "Perform OCR on this scanned PDF document and extract all text verbatim. Preserve formatting, headings, and sections (such as Education, Experience, Projects, Skills, Certifications, Achievements, Languages, Summary) exactly. Do not summarize, explain, or edit the text. Output ONLY the extracted text."
            }
          ]
        });
        extractedText = ocrResponse.text || "";
      } catch (ocrErr: any) {
        console.error("Gemini OCR fallback failed:", ocrErr.message || ocrErr);
      }
    }

    // Backend validation: if text is completely missing, reject (Rule 15)
    if (!extractedText || extractedText.trim().length < 20) {
      return res.status(400).json({
        success: false,
        error: "Failed backend validation: Unable to extract any readable text from the PDF resume."
      });
    }

    // 4. Send ONLY the extracted text to Gemini (Rule 4)
    const prompt = `
You are a highly precise, production-grade AI resume parsing assistant.
Below is the extracted raw text of a candidate's resume, and a target Job Description.

TARGET JOB DESCRIPTION:
"${jobDescription}"

RAW RESUME TEXT:
"""
${extractedText}
"""

YOUR TASK:
Extract candidate profile data strictly from the raw resume text and output valid JSON matching the schema below.

STRICT INSTRUCTIONS:
1. NEVER HALLUCINATE OR FABRICATE ANY DATA. If any information is not explicitly and clearly mentioned in the RAW RESUME TEXT, you MUST return null for that field. Never infer, guess, or extrapolate.
2. Email & Phone:
   - Extract only the genuine, explicitly stated email and phone number.
   - Ignore general web links or non-candidate contact info.
   - If not found or if incomplete/invalid, return null.
3. Ignore random URLs. Under 'linkedin', 'github', and 'portfolio', only include genuine personal candidate URLs. Return null if not explicitly present.
4. Education (Chronological Order):
   - Extract 'institute' (university/school name), 'degree', 'branch' (major), 'cgpa' (GPA or CGPA value), 'start' (start date/year), 'end' (end date/year).
   - Never invent degrees, branches, dates, or CGPA. If missing in the resume text, set to null.
5. Experience (Chronological Order):
   - Extract 'company', 'role' (job title), 'duration' (exact dates or length), and 'description' (responsibilities).
   - Do not invent experience or roles.
6. Projects:
   - Extract project titles exactly.
   - Extract 'technologies' listed explicitly. Do not assume or invent technologies.
7. Skills:
   - Categorize skills strictly into 'languages', 'frameworks', 'tools', 'cloud', 'databases'.
   - Do not invent skills.
8. Add a "confidence_scores" object indicating how clearly and explicitly each field was present in the resume. This must be an integer between 0 and 100 representing your confidence level (e.g., if explicitly listed, confidence is 100; if completely missing, confidence is 0).
   Rate these fields: "email", "phone", "education", "experience", "projects", "skills", "overall".

EXPECTED JSON FORMAT:
{
  "name": "Candidate Name (or null if not found)",
  "email": "Candidate Email (or null if not found)",
  "phone": "Candidate Phone Number (or null if not found)",
  "location": "Candidate Location (or null if not found)",
  "linkedin": "LinkedIn URL (or null if not found)",
  "github": "GitHub URL (or null if not found)",
  "portfolio": "Portfolio URL (or null if not found)",
  "education": [
    {
      "institute": "University/Institution Name (or null)",
      "degree": "Degree (or null)",
      "branch": "Branch/Major (or null)",
      "cgpa": "CGPA/GPA (or null)",
      "start": "Start Date (or null)",
      "end": "End Date (or null)"
    }
  ],
  "experience": [
    {
      "company": "Company Name (or null)",
      "role": "Role/Title (or null)",
      "duration": "Duration/Dates (or null)",
      "description": "Responsibilities/Scope (or null)"
    }
  ],
  "projects": [
    {
      "title": "Project Title (or null)",
      "description": "Project Description (or null)",
      "technologies": ["Technology Name"]
    }
  ],
  "skills": {
    "languages": ["Languages"],
    "frameworks": ["Frameworks"],
    "tools": ["Tools"],
    "cloud": ["Cloud"],
    "databases": ["Databases"]
  },
  "certifications": ["Certifications or empty array"],
  "achievements": ["Achievements or empty array"],
  "soft_skills": ["Soft Skills or empty array"],
  "summary": "Professional 2-sentence summary based ONLY on verified extracted details (or null)",
  "recommendedRole": "Best matched professional title for this candidate based on the resume (or null)",
  "strengths": ["Strengths (at least 3 core professional strengths relative to JD, or empty array)"],
  "weaknesses": ["Areas of improvement (at least 2 concrete professional weaknesses relative to JD, or empty array)"],
  "leadership": 80, (integer between 0-100 indicating leadership indicators based ONLY on actual details, default 50),
  "communication": 80, (integer between 0-100 indicating communication/collaboration based ONLY on actual details, default 50),
  "learningAbility": 80, (integer between 0-100 indicating learning indicators based ONLY on actual details, default 50),
  "total_experience_years": 5, (estimated years of experience as an integer based strictly on experience duration),
  "overallScore": 85, (match score out of 100 relative to the Job Description, calculated ONLY after successful parsing),
  "matchReason": "Detailed explanation of why they fit or do not fit the JD, and how they contribute to a cohesive team.",
  "confidence_scores": {
    "email": 100,
    "phone": 100,
    "education": 100,
    "experience": 100,
    "projects": 100,
    "skills": 100,
    "overall": 100
  }
}

Return ONLY valid JSON. Do not include markdown code block syntax (like \`\`\`json).
`;

    const response = await aiClient.models.generateContent({
      model: "gemini-3.5-flash",
      contents: [{ text: prompt }],
      config: {
        responseMimeType: "application/json",
      }
    });

    const parsedText = response.text || "{}";
    let candidateData: any;
    try {
      candidateData = JSON.parse(parsedText);
    } catch (parseErr) {
      const cleaned = parsedText.replace(/```json/g, "").replace(/```/g, "").trim();
      candidateData = JSON.parse(cleaned);
    }

    // Helper functions for backend validations (Rule 6, 15)
    const isValidEmail = (email: any): boolean => {
      if (typeof email !== "string") return false;
      const re = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
      return re.test(email.trim());
    };

    const isValidPhone = (phone: any): boolean => {
      if (typeof phone !== "string") return false;
      const digits = phone.replace(/\D/g, "");
      return digits.length >= 7 && digits.length <= 20;
    };

    // Robust backend validation of critical fields (Rule 15)
    if (!candidateData || !candidateData.name || candidateData.name.toLowerCase().includes("null")) {
      candidateData.name = originalName.replace(".pdf", "").replace(/[-_]/g, " ").replace(/\b\w/g, c => c.toUpperCase());
    }

    // Validation & cleaning of contact details (Rule 6, 7)
    if (candidateData.email) {
      if (!isValidEmail(candidateData.email)) {
        candidateData.email = null;
      }
    } else {
      candidateData.email = null;
    }

    if (candidateData.phone) {
      if (!isValidPhone(candidateData.phone)) {
        candidateData.phone = null;
      }
    } else {
      candidateData.phone = null;
    }

    // Prepend warning if low overall confidence (Rule 12)
    const overallConfidence = Number(candidateData.confidence_scores?.overall) || 100;
    let finalMatchReason = candidateData.matchReason || "Parsed successfully.";

    if (overallConfidence < 90) {
      finalMatchReason = `Low confidence. Manual review recommended.\n\n${finalMatchReason}`;
    }

    // Append confidence score summary block (Rule 14)
    const emailConf = candidateData.confidence_scores?.email !== undefined ? `${candidateData.confidence_scores.email}%` : "N/A";
    const eduConf = candidateData.confidence_scores?.education !== undefined ? `${candidateData.confidence_scores.education}%` : "N/A";
    const projConf = candidateData.confidence_scores?.projects !== undefined ? `${candidateData.confidence_scores.projects}%` : "N/A";
    const skillConf = candidateData.confidence_scores?.skills !== undefined ? `${candidateData.confidence_scores.skills}%` : "N/A";
    const expConf = candidateData.confidence_scores?.experience !== undefined ? `${candidateData.confidence_scores.experience}%` : "N/A";

    finalMatchReason += `\n\n[Extraction Confidence Scores]\n- Email: ${emailConf}\n- Education: ${eduConf}\n- Projects: ${projConf}\n- Skills: ${skillConf}\n- Experience: ${expConf}`;

    // Flatten categorized skills list (Rule 11)
    let flatSkills: string[] = [];
    if (candidateData.skills) {
      const s = candidateData.skills;
      flatSkills = [
        ...(s.languages || []),
        ...(s.frameworks || []),
        ...(s.tools || []),
        ...(s.cloud || []),
        ...(s.databases || [])
      ].filter((val: any) => typeof val === "string" && val.trim().length > 0);
    }

    // Format education summary for UI (Rule 8)
    let eduSummary = "No formal education listed";
    if (candidateData.education && Array.isArray(candidateData.education) && candidateData.education.length > 0) {
      const edu = candidateData.education[0];
      const parts = [
        edu.degree,
        edu.branch,
        edu.institute,
        edu.cgpa ? `CGPA: ${edu.cgpa}` : null
      ].filter((val: any) => val && typeof val === "string" && !val.toLowerCase().includes("null") && !val.toLowerCase().includes("unknown"));
      if (parts.length > 0) {
        eduSummary = parts.join(", ");
      }
    }

    // Map projects (Rule 10)
    const mappedProjects = (candidateData.projects || []).map((p: any) => ({
      name: p.title || "Untitled Project",
      description: p.description || "No description provided.",
      technologies: p.technologies || []
    }));

    // Create verified candidate
    const newId = `cand-${Date.now()}`;
    const newCandidate: Candidate = {
      id: newId,
      name: candidateData.name,
      skills: flatSkills,
      experience: Number(candidateData.total_experience_years) || 1,
      projects: mappedProjects,
      certifications: candidateData.certifications || [],
      education: eduSummary,
      leadership: Number(candidateData.leadership) || 50,
      communication: Number(candidateData.communication) || 50,
      learningAbility: Number(candidateData.learningAbility) || 50,
      summary: candidateData.summary || "No summary extracted.",
      recommendedRole: candidateData.recommendedRole || "Developer",
      strengths: candidateData.strengths || [],
      weaknesses: candidateData.weaknesses || [],
      overallScore: Number(candidateData.overallScore) || 70,
      matchReason: finalMatchReason,
      email: candidateData.email || undefined,
      phone: candidateData.phone || undefined,
      achievements: candidateData.achievements || []
    };

    // Save strictly validated candidate to memory database (Rule 15)
    candidatesDb.push(newCandidate);
    return res.json({
      success: true,
      candidate: newCandidate,
      message: `Resume parsed and matched successfully via Gemini with ${parseMethod} extraction!`
    });

  } catch (error: any) {
    console.error("Resume Parse Error:", error);
    res.status(500).json({ success: false, error: error.message || "Failed to parse PDF resume." });
  }
});

// API: Manual Add Candidate (for ease of hackathon demo)
app.post("/api/candidates", (req, res) => {
  const newCandidate: Candidate = {
    id: `cand-${Date.now()}`,
    ...req.body
  };
  candidatesDb.push(newCandidate);
  res.json({ success: true, candidate: newCandidate });
});

// API: Edit Candidate
app.put("/api/candidates/:id", (req, res) => {
  const index = candidatesDb.findIndex(c => c.id === req.params.id);
  if (index !== -1) {
    candidatesDb[index] = { ...candidatesDb[index], ...req.body };
    res.json({ success: true, candidate: candidatesDb[index] });
  } else {
    res.status(404).json({ success: false, error: "Candidate not found" });
  }
});

// API: Delete Candidate
app.delete("/api/candidates/:id", (req, res) => {
  const initialLength = candidatesDb.length;
  candidatesDb = candidatesDb.filter(c => c.id !== req.params.id);
  if (candidatesDb.length < initialLength) {
    res.json({ success: true, message: "Candidate deleted successfully" });
  } else {
    res.status(404).json({ success: false, error: "Candidate not found" });
  }
});

// API: AI Team Builder
app.post("/api/teambuilder", async (req, res) => {
  try {
    const { requirements, teamSize, whatIfExcludeIds, whatIfIncludeIds } = req.body;
    const size = Number(teamSize) || 4;
    const reqText = requirements || "Need 1 Backend, 1 Frontend, 1 AI Engineer, 1 UI Designer";

    // Handle What-If simulations
    const excludeIds: string[] = Array.isArray(whatIfExcludeIds) ? whatIfExcludeIds : [];
    const filteredCandidates = candidatesDb.filter(c => !excludeIds.includes(c.id));

    // Compile simulated impact statement
    let whatIfAnalysis = {
      impactScore: 100,
      impactStatement: "Operating on full capacity talent pool. Stable and recommended.",
      riskAlerts: [] as string[]
    };

    if (excludeIds.length > 0) {
      const excludedNames = candidatesDb.filter(c => excludeIds.includes(c.id)).map(c => c.name);
      let scoreReduction = excludeIds.length * 8;
      const alerts: string[] = [];
      let statement = `Simulating team compilation with ${excludedNames.join(", ")} excluded. `;

      if (excludeIds.includes("cand-2")) {
        statement += "Excluding David Chen (Principal Systems Architect) reduces high-throughput backend infrastructure stability by 40% and raises latency risk under load. ";
        alerts.push("CRITICAL: Loss of deep backend and container orchestration expertise.");
        scoreReduction += 12;
      }
      if (excludeIds.includes("cand-3")) {
        statement += "Excluding Dr. Elena Rostova removes Stanford Ph.D. machine learning research oversight, lowering the innovation score from 96% to 62%. ";
        alerts.push("WARNING: Loss of specialized Generative AI RAG capabilities.");
        scoreReduction += 10;
      }
      if (excludeIds.includes("cand-1")) {
        statement += "Excluding Sarah Jenkins eliminates core Next.js and frontend design system governance, raising client-side delivery friction. ";
        alerts.push("WARNING: Design system governance is unassigned.");
      }
      if (excludeIds.includes("cand-5")) {
        statement += "Excluding Rajesh Kumar removes our Certified ScrumMaster, impacting agile sprint velocity and cross-functional team coordination. ";
        alerts.push("WARNING: Agile Scrum coordination risk.");
      }

      whatIfAnalysis = {
        impactScore: Math.max(40, 100 - scoreReduction),
        impactStatement: statement.trim() || `Candidate capacity reduced. Team parameters re-evaluated with the remaining ${filteredCandidates.length} profiles.`,
        riskAlerts: alerts
      };
    }

    // Try Gemini if client is initialized
    if (aiClient) {
      try {
        const prompt = `
          You are an expert Organizational Psychologist and Agile Engineering Director.
          We have the following list of available candidates:
          ${JSON.stringify(filteredCandidates, null, 2)}

          The team requirement is:
          "${reqText}"
          Target team size: ${size}

          You MUST construct 5 different team archetypes from this candidate pool:
          1. "best": The absolute optimal team based on overall skills, leadership, and compatibility.
          - "alternative": An alternative team option that does NOT include the highest-scoring candidate from the "best" team, to give recruiters choices.
          - "budget": A cost-effective team emphasizing high-learning, high-efficiency candidates with slightly less senior years overhead.
          - "innovation": A high-innovation R&D squad biased towards research, design, and machine learning specialists.
          - "delivery": A high-throughput delivery squad biased towards high-experience, backend, systems, and scrum leadership.

          For EACH of these 5 archetypes, calculate quantitative compatibility scores (integers between 0 and 100), including:
          - skillDiversity: coverage of unique skills
          - leadershipBalance: mix of strategic leaders vs doers
          - experienceBalance: distribution of senior/mid levels
          - communication: communication average adjusted for chemistry
          - learningPotential: learning potential average
          - conflictRisk: (0 is lowest, 100 is highest. Keep low for good chemistry)
          - innovationScore: design and AI capacity
          - deliveryConfidence: execution speed
          - overall: consolidated compatibility score
          - burnoutRisk: burnout risk score (integer)
          - knowledgeCoverage: knowledge coverage score (integer)
          - communicationBalance: balance of communication patterns (integer)

          Also include a "teamDna" object for each archetype:
          {
            "innovationIndex": number,
            "executionIndex": number,
            "leadershipIndex": number,
            "learningSpeed": number,
            "communicationIndex": number,
            "technicalCoverage": number,
            "riskLevel": "Low" | "Medium" | "High",
            "recommendedLead": "Name of recommended candidate to lead",
            "workingStyle": "e.g. Hybrid Agile Scrum"
          }

          Generate the output as a structured JSON object with EXACTLY these five keys:
          {
            "best": { ...TeamComposition },
            "alternative": { ...TeamComposition },
            "budget": { ...TeamComposition },
            "innovation": { ...TeamComposition },
            "delivery": { ...TeamComposition }
          }

          Return ONLY valid JSON conforming strictly to this format. Do not write markdown tags.
        `;

        const response = await aiClient.models.generateContent({
          model: "gemini-3.5-flash",
          contents: prompt,
          config: {
            responseMimeType: "application/json",
          }
        });

        const parsedText = response.text || "{}";
        let teamArchetypes;
        try {
          teamArchetypes = JSON.parse(parsedText);
        } catch (parseErr) {
          const cleaned = parsedText.replace(/```json/g, "").replace(/```/g, "").trim();
          teamArchetypes = JSON.parse(cleaned);
        }

        // Validate structure has the key archetypes, if missing some, merge with fallback
        if (teamArchetypes && teamArchetypes.best) {
          return res.json({
            success: true,
            data: teamArchetypes,
            whatIfAnalysis,
            fallbackUsed: false
          });
        }
      } catch (geminiErr: any) {
        console.warn("Gemini team builder failed. Activating robust local fallback:", geminiErr.message || geminiErr);
      }
    }

    // Fallback mode: generates all 5 archetypes with high fidelity
    const fallbackArchetypes = buildTeamFallback(filteredCandidates, reqText, size);
    res.json({
      success: true,
      data: fallbackArchetypes,
      whatIfAnalysis,
      fallbackUsed: true
    });

  } catch (error: any) {
    console.error("Team Builder Error:", error);
    res.status(500).json({ success: false, error: error.message || "Failed to generate AI team recommendation." });
  }
});

// API: Candidate Detailed Analysis (Skill Gap & Interview Questions)
app.post("/api/candidates/:id/analysis", async (req, res) => {
  try {
    const candidate = candidatesDb.find(c => c.id === req.params.id);
    if (!candidate) {
      return res.status(404).json({ success: false, error: "Candidate not found" });
    }

    const { targetRole, jobDescription } = req.body;
    const role = targetRole || candidate.recommendedRole;
    const jd = jobDescription || "General Software Developer with AI integrations.";

    // Try Gemini if client is initialized
    if (aiClient) {
      try {
        const prompt = `
          You are an expert Technical Interviewer and Skill Development Consultant.
          Analyze this candidate for the following target position:
          
          CANDIDATE:
          ${JSON.stringify(candidate, null, 2)}
          
          TARGET ROLE:
          "${role}"
          
          JOB DESCRIPTION SUMMARY:
          "${jd}"

          Perform a comprehensive gap analysis and interview roadmap:
          1. Overall Score out of 100 for this specific role.
          2. Strengths and weaknesses relative to this role.
          3. Skill Gap Analysis: Identify missing or partial skills needed. For each gap, generate a 2-step customized learning roadmap (Resource/Class, estimatedDays, and detailed description).
          4. Interview Question Generator: Create 6 high-quality custom questions tailored specifically to this candidate's history and gaps:
             - 2 Technical Questions (probing backend/frontend/AI limitations)
             - 2 Behavioral Questions (probing cooperation and leadership scores)
             - 2 Scenario/Situational Questions (presenting realistic project problems relevant to their history)
             Provide the question, the expected perfect answer, and difficulty (Easy, Medium, Hard).

          Return the analysis in structured JSON matching this exact layout:
          {
            "candidateId": "${candidate.id}",
            "candidateName": "${candidate.name}",
            "score": 88,
            "roleAlignment": "Perfect alignment for core React development, with minor server-side training requirements.",
            "strengths": ["Strength 1", "Strength 2"],
            "weaknesses": ["Improvement area 1", "Improvement area 2"],
            "skillGaps": [
              {
                "skill": "Name of missing or partial skill",
                "status": "missing", (or "partial" or "acquired"),
                "roadmap": [
                  {
                    "resource": "Specific online resource, documentation, or certification name",
                    "estimatedDays": 14,
                    "description": "Exactly what topic to study and build to master this gap"
                  }
                ]
              }
            ],
            "interviewQuestions": [
              {
                "type": "technical",
                "question": "The interview question details",
                "expectedAnswer": "Brief bullet points of what a great candidate response contains",
                "difficulty": "Medium"
              }
            ]
          }

          Return ONLY valid JSON.
        `;

        const response = await aiClient.models.generateContent({
          model: "gemini-3.5-flash",
          contents: prompt,
          config: {
            responseMimeType: "application/json",
          }
        });

        const parsedText = response.text || "{}";
        let analysisData;
        try {
          analysisData = JSON.parse(parsedText);
        } catch (parseErr) {
          const cleaned = parsedText.replace(/```json/g, "").replace(/```/g, "").trim();
          analysisData = JSON.parse(cleaned);
        }

        return res.json({ success: true, data: analysisData, fallbackUsed: false });
      } catch (geminiErr: any) {
        console.warn("Gemini candidate detailed analysis failed. Activating robust local fallback:", geminiErr.message || geminiErr);
      }
    }

    // Fallback mode
    const fallbackAnalysis = analyzeCandidateFallback(candidate, role, jd);
    res.json({ success: true, data: fallbackAnalysis, fallbackUsed: true });

  } catch (error: any) {
    console.error("Analysis Error:", error);
    res.status(500).json({ success: false, error: error.message || "Failed to generate detailed analysis." });
  }
});

// API: Natural Language Recruiter Chat Assistant
app.post("/api/chat", async (req, res) => {
  try {
    const { messages, query } = req.body;
    const recentQuery = query || (messages && messages[messages.length - 1]?.text) || "";

    // Try Gemini if client is initialized
    if (aiClient) {
      try {
        const systemPrompt = `
          You are "FusionAI", the interactive Natural Language AI Assistant of "TalentFusion AI" Workforce Intelligence platform.
          Your job is to help recruiters search, evaluate, filter, and structure candidates or build specialized teams.
          You have access to the current active Candidate Database:
          ${JSON.stringify(candidatesDb, null, 2)}

          Analyze the recruiter's prompt. You can understand requests like:
          - "Find candidates with React and Go"
          - "Who is best suited for AI?"
          - "Identify candidate strengths"
          - "Show candidates suitable for Team Lead"
          - "Build the best team"

          Synthesize a helpful, conversational, and expert HR response.
          Additionally, you can specify relevant matching candidate IDs in the "candidateIds" array if specific candidates are discussed, matched, or searched for.
          You can also trigger UI interactions by returning a "suggestedAction" object.
          
          Trigger Actions guide:
          - If they ask for specific candidates or filters: suggest action: { "type": "view_candidate", "payload": { "id": "ID_OF_BEST_MATCH_CANDIDATE" } }
          - If they want to build a team or select teams: suggest action: { "type": "build_team", "payload": { "requirements": "Extracted requirements description" } }
          - If they want to look at the Talent Graph: suggest action: { "type": "view_graph", "payload": {} }
          - If they want gap analysis: suggest action: { "type": "show_gap", "payload": { "id": "ID_OF_CANDIDATE" } }

          Response schema to return:
          {
            "text": "Your professional conversational reply here. Use clean HTML formatting like <b>bolding</b> or lists if needed.",
            "candidateIds": ["cand-1", "cand-2"], (OPTIONAL - array of candidate IDs that are relevant or match the search terms/discussion)
            "suggestedAction": { "type": "view_candidate", "payload": { "id": "cand-1" } } (OPTIONAL - only return if relevant to trigger action)
          }

          Return ONLY valid JSON.
        `;

        const chatHistory = messages ? messages.map((m: any) => ({
          role: m.sender === "user" ? "user" : "model",
          parts: [{ text: m.text }]
        })) : [];

        // Add recent query
        chatHistory.push({
          role: "user",
          parts: [{ text: recentQuery }]
        });

        const response = await aiClient.models.generateContent({
          model: "gemini-2.5-flash",
          contents: chatHistory.length > 0 ? [
            { text: systemPrompt },
            ...chatHistory.slice(-10) // Send last 10 messages for context
          ] : [
            { text: systemPrompt },
            { text: recentQuery }
          ],
          config: {
            responseMimeType: "application/json",
          }
        });

        const parsedText = response.text || "{}";
        let chatReply;
        try {
          chatReply = JSON.parse(parsedText);
        } catch (parseErr) {
          const cleaned = parsedText.replace(/```json/g, "").replace(/```/g, "").trim();
          chatReply = JSON.parse(cleaned);
        }

        return res.json({ success: true, data: chatReply, fallbackUsed: false });
      } catch (geminiErr: any) {
        console.warn("Gemini chat interaction failed. Activating robust local fallback:", geminiErr.message || geminiErr);
      }
    }

    // Fallback mode
    const fallbackReply = chatFallback(messages || [], recentQuery);
    res.json({ success: true, data: fallbackReply, fallbackUsed: true });

  } catch (error: any) {
    console.error("Chat Error:", error);
    res.status(500).json({ success: false, error: error.message || "Failed to process chat query." });
  }
});

// API: Excel Export of Candidates Database
app.get("/api/export/excel", (req, res) => {
  try {
    const dataRows = candidatesDb.map((c, index) => {
      return {
        "Rank": index + 1,
        "Candidate Name": c.name,
        "Recommended Role": c.recommendedRole,
        "Overall Score": c.overallScore,
        "Skill Match %": Math.round(c.overallScore * 0.95), // realistic skill match ratio
        "Leadership (0-100)": c.leadership,
        "Communication (0-100)": c.communication,
        "Learning Ability (0-100)": c.learningAbility,
        "Strengths": c.strengths.join(", "),
        "Weaknesses": c.weaknesses.join(", "),
        "Recommendation": c.overallScore >= 90 ? "Strong Hire" : (c.overallScore >= 80 ? "Hire" : "Review Candidate")
      };
    });

    // Create workbook
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(dataRows);

    // Styling column widths roughly
    ws["!cols"] = [
      { wch: 6 },  // Rank
      { wch: 22 }, // Name
      { wch: 30 }, // Role
      { wch: 14 }, // Score
      { wch: 14 }, // Skill Match %
      { wch: 12 }, // Leadership
      { wch: 14 }, // Communication
      { wch: 14 }, // Learning
      { wch: 45 }, // Strengths
      { wch: 45 }, // Weaknesses
      { wch: 18 }  // Recommendation
    ];

    XLSX.utils.book_append_sheet(wb, ws, "TalentFusion Registry");

    // Write buffer
    const buffer = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });

    res.setHeader("Content-Disposition", 'attachment; filename="TalentFusion_Workforce_Report.xlsx"');
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.send(buffer);
  } catch (error: any) {
    console.error("Excel Export Error:", error);
    res.status(500).json({ success: false, error: "Failed to generate Excel report." });
  }
});

// Vite Middleware integration for Full Stack Development / Asset Serving
async function startServer() {
  // Mount Vite development server middleware if not running in production
  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    // Serve static frontend assets from dist in production
    const distPath = path.join(process.cwd(), "dist");
    app.use(express.static(distPath));
    app.get("*", (req, res) => {
      res.sendFile(path.join(distPath, "index.html"));
    });
  }

  // Handle errors
  app.use((err: any, req: express.Request, res: express.Response, next: express.NextFunction) => {
    console.error("Unhandled Server Error:", err);
    res.status(500).json({ success: false, error: err.message || "An unexpected error occurred." });
  });

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`[TalentFusion AI] Server running on http://localhost:${PORT}`);
  });
}

startServer();
