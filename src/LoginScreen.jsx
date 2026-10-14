import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { QUOTES, QUOTE_AUTHOR } from "./quotes.js";
import { supabase, normalizeEmail, isWorkEmail, siteAddress } from "./supabase.js";

// Sign in screen for the CP2 Gate Log: MMR mark, an industrial skyline drawn
// in MMR colors (no site photos, so nothing restricted ends up on a public
// page), the account forms, and a rotating panel of Pepper Rutland quotes.
//
// Accounts: anyone with an @mmrgrp.com email can create one. They get an email
// link to confirm the address, then choose their own password. The server
// enforces the domain rule, and the database only answers to confirmed accounts.

const LOGO_RATIO = 466.2 / 227.8;

// MMR mark, paths taken unchanged from MMR's official white logo artwork.
function MmrLogo({ height = 52 }) {
  return (
    <svg
      viewBox="0 0 466.2 227.8"
      width={Math.round(height * LOGO_RATIO)}
      height={height}
      role="img"
      aria-label="MMR"
      style={{ display: "block" }}
    >
      <path fill="currentColor" d="M451.7,182.2V30.8c0-15.4-15.4-30.8-30.7-30.8H30.7C15.4,0,0,15.4,0,30.8v166.3c0,15.4,15.4,30.7,30.7,30.7h378.8s10.1.1,12.3-.6c-5.7-1.9-18.3-7.9-30.8-30.1H35.2l-8,10c-1.2,1.4-3.2,1.6-4.6.4-1.4-1.1-1.7-3.2-.6-4.6l8.8-11V36.7l-8.8-11c-1.1-1.4-.8-3.5.6-4.6,1.4-1.1,3.5-.9,4.6.4l7.5,9.3h382.6l7.5-9.3c1.2-1.4,3.2-1.6,4.6-.4,1.4,1.1,1.7,3.2.6,4.6l-9,11.3v126.1l8.6,14.9c.9,1.6.4,3.7-1.2,4.7-1.6,1-3.7.5-4.7-1.1l-2.2-3.6c-.1-.2-28.9-55.8-28.9-55.8,1.4-1.6,16.5-10.7,16.5-37.3s-15.5-41.2-31.7-41.2h-48.4c-4,0-7.1,2.3-7.1,6.3v134.6h29.7v-105.4c0-3.9,3.1-6.3,7.1-6.3h9.5c5.5,0,8.3,3.4,8.3,11.6s-11.4,20.9-18.1,27.5c-4.3,4.3-.6,7.9,13.8,37.4,7.6,14.2,20.5,34.5,25.5,42,8.3,12.2,15.7,21.4,21.8,25.2,18.9,10.8,32.1-13.3,32.1-34.2" />
      <path fill="currentColor" d="M458.4,215.1c0-1.5-1.1-2.1-2.6-2.1s-1.2,0-1.4.1v4.2h1.5c1.6,0,2.5-.9,2.5-2.2M459.4,220.2c.4,1.6.6,2.7.9,3.1h-1.5c-.2-.3-.4-1.3-.7-2.7-.3-1.6-1-2.2-2.3-2.2h-1.4v5h-1.5v-11.3c.7-.2,1.8-.3,2.8-.3,1.6,0,2.6.3,3.3.9.6.5.9,1.3.9,2.2,0,1.5-1,2.5-2.1,2.9h0c.9.4,1.4,1.2,1.7,2.4M463.4,213.3c-.8-1.4-1.8-2.4-3.1-3.2-1.3-.7-2.7-1.1-4.1-1.1s-2.8.4-4.1,1.1c-1.3.7-2.4,1.8-3.1,3.2-.8,1.4-1.1,2.8-1.1,4.3s.4,2.9,1.1,4.2c.7,1.4,1.8,2.4,3.1,3.2,1.3.8,2.7,1.1,4.2,1.1s2.9-.4,4.2-1.1c1.3-.8,2.4-1.8,3.1-3.2.7-1.4,1.1-2.8,1.1-4.2s-.4-2.9-1.1-4.3M464.9,222.7c-.9,1.6-2.1,2.9-3.7,3.8-1.6.9-3.3,1.3-5,1.3s-3.4-.4-5-1.3c-1.6-.9-2.8-2.2-3.7-3.8-.9-1.6-1.3-3.3-1.3-5.1s.5-3.5,1.4-5.1c.9-1.6,2.2-2.9,3.8-3.8,1.6-.9,3.3-1.3,5-1.3s3.4.4,5,1.3c1.6.9,2.9,2.1,3.8,3.8.9,1.6,1.4,3.4,1.4,5.1s-.4,3.5-1.3,5.1" />
      <path fill="currentColor" d="M171.6,184.5V49.9c0-3.9-3.1-6.3-7.1-6.3h-18.6c-4.9,0-6.2.8-7.6,3.6-1.1,2.3-23.7,51-26.3,56.7l5.6,12.1c.7,1.6,0,3.4-1.6,4.1-1.6.7-3.4,0-4.2-1.5,0,0-33.4-71.1-33.5-71.4-1.4-2.9-2.7-3.6-7.6-3.6h-20c-4,0-7.1,2.3-7.1,6.3v134.6h30.7v-64.8s27.8,53,30,56.7c2.1,3.7,3.9,3.7,6.1,0,2.4-3.7,31.8-56.8,31.8-56.8v64.8h29.5Z" />
      <path fill="currentColor" d="M310.9,184.5V49.9c0-3.9-3.1-6.3-7.1-6.3h-18.6c-4.9,0-6.2.8-7.6,3.6-1.1,2.3-23.7,51-26.3,56.7l5.6,12.1c.7,1.6,0,3.4-1.6,4.1-1.6.7-3.4,0-4.2-1.5,0,0-33.4-71.1-33.5-71.4-1.4-2.9-2.7-3.6-7.6-3.6h-20c-4,0-7.1,2.3-7.1,6.3v134.6h30.7v-64.8s27.8,53,30,56.7c2.1,3.7,3.9,3.7,6.1,0,2.4-3.7,31.8-56.8,31.8-56.8v64.8h29.5Z" />
    </svg>
  );
}

// Lattice boom for the crawler crane, built from two chords and zigzag bracing.
function boomLattice(ax, ay, bx, by, width, segments) {
  const dx = bx - ax, dy = by - ay;
  const len = Math.hypot(dx, dy);
  const nx = (-dy / len) * width, ny = (dx / len) * width;
  const lines = [];
  for (let i = 0; i < segments; i++) {
    const t0 = i / segments, t1 = (i + 1) / segments;
    const x0 = ax + dx * t0, y0 = ay + dy * t0;
    const x1 = ax + dx * t1, y1 = ay + dy * t1;
    lines.push(i % 2 === 0
      ? [x0, y0, x1 + nx, y1 + ny]
      : [x0 + nx, y0 + ny, x1, y1]);
  }
  return { nx, ny, lines };
}

function useNarrowScreen() {
  const query = "(max-width: 640px)";
  const [narrow, setNarrow] = useState(() =>
    typeof window !== "undefined" && !!window.matchMedia && window.matchMedia(query).matches);
  useEffect(() => {
    if (!window.matchMedia) return undefined;
    const mq = window.matchMedia(query);
    const onChange = () => setNarrow(mq.matches);
    if (mq.addEventListener) mq.addEventListener("change", onChange);
    else mq.addListener(onChange);
    return () => {
      if (mq.removeEventListener) mq.removeEventListener("change", onChange);
      else mq.removeListener(onChange);
    };
  }, []);
  return narrow;
}

function SiteBackdrop() {
  const narrow = useNarrowScreen();
  const boom = boomLattice(252, 450, 560, 122, 9, 16);
  const bents = Array.from({ length: 26 }, (_, i) => i * 64);
  return (
    <div className="lg-backdrop" aria-hidden="true">
      <svg className="lg-grid" width="100%" height="100%">
        <defs>
          <pattern id="lgGridMinor" width="32" height="32" patternUnits="userSpaceOnUse">
            <path d="M32 0H0V32" fill="none" stroke="rgba(255,255,255,0.035)" strokeWidth="1" />
          </pattern>
          <pattern id="lgGridMajor" width="160" height="160" patternUnits="userSpaceOnUse">
            <rect width="160" height="160" fill="url(#lgGridMinor)" />
            <path d="M160 0H0V160" fill="none" stroke="rgba(255,255,255,0.065)" strokeWidth="1" />
          </pattern>
        </defs>
        <rect width="100%" height="100%" fill="url(#lgGridMajor)" />
      </svg>

      <svg className="lg-skyline" viewBox="0 0 1600 520" preserveAspectRatio={narrow ? "xMaxYMax slice" : "xMidYMax slice"}>
        <defs>
          <linearGradient id="lgTank" x1="0" y1="0" x2="0" y2="1">
            <stop offset="0" stopColor="#2c1117" />
            <stop offset="1" stopColor="#14090b" />
          </linearGradient>
          <linearGradient id="lgFlame" x1="0" y1="1" x2="0" y2="0">
            <stop offset="0" stopColor="#f8c700" />
            <stop offset="0.6" stopColor="#ffd84a" />
            <stop offset="1" stopColor="#fff4c2" />
          </linearGradient>
          <radialGradient id="lgFlameGlow">
            <stop offset="0" stopColor="rgba(248,199,0,0.55)" />
            <stop offset="0.45" stopColor="rgba(248,199,0,0.14)" />
            <stop offset="1" stopColor="rgba(248,199,0,0)" />
          </radialGradient>
        </defs>

        {/* Distant process trains */}
        <g fill="#2b1319" opacity="0.75">
          <rect x="90" y="250" width="24" height="260" rx="8" />
          <rect x="126" y="300" width="16" height="210" rx="6" />
          <rect x="150" y="400" width="120" height="110" />
          <rect x="300" y="228" width="30" height="282" rx="10" />
          <rect x="340" y="290" width="18" height="220" rx="6" />
          <rect x="362" y="380" width="92" height="130" />
          <rect x="996" y="258" width="26" height="252" rx="8" />
          <rect x="1036" y="318" width="16" height="192" rx="6" />
          <rect x="1392" y="388" width="82" height="122" />
          <rect x="1480" y="238" width="28" height="272" rx="9" />
          <rect x="1520" y="300" width="18" height="210" rx="6" />
        </g>
        <g fill="#f8c700" opacity="0.8">
          <circle cx="102" cy="262" r="1.8" />
          <circle cx="315" cy="240" r="1.8" />
          <circle cx="1009" cy="270" r="1.8" />
          <circle cx="1494" cy="250" r="1.8" />
          <circle cx="210" cy="412" r="1.4" />
          <circle cx="408" cy="392" r="1.4" />
          <circle cx="1432" cy="400" r="1.4" />
        </g>

        {/* LNG storage tanks */}
        <path d="M520 510 V350 Q690 300 860 350 V510 Z" fill="url(#lgTank)" />
        <path d="M520 350 Q690 300 860 350" fill="none" stroke="rgba(248,199,0,0.13)" strokeWidth="1.5" />
        <line x1="540" y1="498" x2="840" y2="352" stroke="rgba(255,255,255,0.03)" strokeWidth="1.5" />
        <rect x="684" y="318" width="12" height="5" fill="#2c1117" />
        <path d="M880 510 V366 Q1020 324 1160 366 V510 Z" fill="url(#lgTank)" />
        <path d="M880 366 Q1020 324 1160 366" fill="none" stroke="rgba(248,199,0,0.11)" strokeWidth="1.5" />
        <line x1="900" y1="498" x2="1140" y2="368" stroke="rgba(255,255,255,0.025)" strokeWidth="1.5" />

        {/* Flare stack */}
        <circle className="lg-flare-glow" cx="1290" cy="138" r="78" fill="url(#lgFlameGlow)" />
        <line x1="1290" y1="172" x2="1196" y2="500" stroke="rgba(255,255,255,0.05)" strokeWidth="1" />
        <line x1="1290" y1="172" x2="1384" y2="500" stroke="rgba(255,255,255,0.05)" strokeWidth="1" />
        <line x1="1290" y1="262" x2="1226" y2="500" stroke="rgba(255,255,255,0.045)" strokeWidth="1" />
        <line x1="1290" y1="262" x2="1354" y2="500" stroke="rgba(255,255,255,0.045)" strokeWidth="1" />
        <rect x="1286" y="150" width="8" height="360" fill="#160a0c" />
        <path className="lg-flame" d="M1290 108 C1280 124 1280 140 1290 149 C1300 140 1300 124 1290 108 Z" fill="url(#lgFlame)" />
        <circle className="lg-beacon" cx="1290" cy="156" r="2.2" fill="#ff4d4d" />

        {/* Pipe rack running across the site */}
        <g>
          <line x1="0" y1="432" x2="1600" y2="432" stroke="#1d0c10" strokeWidth="5" />
          <line x1="0" y1="440" x2="1600" y2="440" stroke="#1f0d11" strokeWidth="8" />
          <rect x="0" y="446" width="1600" height="5" fill="#170a0d" />
          <rect x="0" y="470" width="1600" height="4" fill="#170a0d" />
          {bents.map(x => (
            <rect key={x} x={x} y="446" width="5" height="60" fill="#13080a" />
          ))}
        </g>

        {/* Crawler crane */}
        <g>
          <line x1="300" y1="332" x2="558" y2="122" stroke="rgba(255,255,255,0.06)" strokeWidth="1" />
          <line x1="232" y1="452" x2="300" y2="332" stroke="#1a0b0e" strokeWidth="3" />
          <line x1="252" y1="450" x2="560" y2="122" stroke="#33151b" strokeWidth="3" />
          <line x1={252 + boom.nx} y1={450 + boom.ny} x2={560 + boom.nx} y2={122 + boom.ny} stroke="#33151b" strokeWidth="3" />
          {boom.lines.map(([x1, y1, x2, y2], i) => (
            <line key={i} x1={x1} y1={y1} x2={x2} y2={y2} stroke="#2c1217" strokeWidth="1.3" />
          ))}
          <line x1="562" y1="126" x2="562" y2="330" stroke="rgba(255,255,255,0.12)" strokeWidth="1" />
          <rect x="556" y="330" width="12" height="16" rx="2" fill="#1d0c10" />
          <rect x="176" y="448" width="96" height="32" fill="#150a0c" />
          <rect x="256" y="436" width="30" height="26" fill="#180b0d" />
          <rect x="263" y="442" width="16" height="10" fill="rgba(248,199,0,0.22)" />
          <rect x="150" y="478" width="150" height="24" rx="7" fill="#110708" />
          <circle className="lg-beacon b2" cx="561" cy="118" r="2.2" fill="#ff4d4d" />
        </g>

        <rect x="0" y="500" width="1600" height="20" fill="#0b0507" />
      </svg>

      <div className="lg-titleblock">
        <div className="lg-tb-row lg-tb-head">Gate access</div>
        <div className="lg-tb-row"><span>Project</span><span>CP2 LNG</span></div>
        <div className="lg-tb-row"><span>Location</span><span>Cameron Parish, LA</span></div>
        <div className="lg-tb-row"><span>System</span><span>Gate Log</span></div>
        <div className="lg-tb-row"><span>Contractor</span><span>MMR Constructors</span></div>
      </div>
    </div>
  );
}

// About 3.5 words a second plus a pause to take it in, between 6 and 16 seconds.
export function readMs(text) {
  const words = String(text || "").trim().split(/\s+/).filter(Boolean).length;
  return Math.min(16000, Math.max(6000, 2500 + words * 290));
}

function shuffledOrder(n, avoidFirst) {
  const a = Array.from({ length: n }, (_, i) => i);
  for (let i = n - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [a[i], a[j]] = [a[j], a[i]];
  }
  if (n > 1 && a[0] === avoidFirst) [a[0], a[1]] = [a[1], a[0]];
  return a;
}

function usePrefersReducedMotion() {
  const query = "(prefers-reduced-motion: reduce)";
  const [reduce, setReduce] = useState(() =>
    typeof window !== "undefined" && !!window.matchMedia && window.matchMedia(query).matches);
  useEffect(() => {
    if (!window.matchMedia) return undefined;
    const mq = window.matchMedia(query);
    const onChange = () => setReduce(mq.matches);
    if (mq.addEventListener) mq.addEventListener("change", onChange);
    else mq.addListener(onChange);
    return () => {
      if (mq.removeEventListener) mq.removeEventListener("change", onChange);
      else mq.removeListener(onChange);
    };
  }, []);
  return reduce;
}

// Shows one quote at a time in random order without repeats until every quote
// has had a turn. Each stays up long enough to read, pauses while hovered or
// while the tab is in the background, and the gold line shows time remaining.
function QuoteRotator({ quotes, author }) {
  const list = useMemo(
    () => (quotes || []).map(q => String(q || "").trim()).filter(Boolean),
    [quotes]
  );
  const reduceMotion = usePrefersReducedMotion();
  const orderRef = useRef(null);
  const posRef = useRef(0);
  const [idx, setIdx] = useState(() => {
    orderRef.current = shuffledOrder(list.length, -1);
    return orderRef.current.length ? orderRef.current[0] : 0;
  });
  const [cycle, setCycle] = useState(0);
  const [shown, setShown] = useState(true);
  const [hoverPause, setHoverPause] = useState(false);
  const [tabHidden, setTabHidden] = useState(
    () => typeof document !== "undefined" && document.visibilityState === "hidden");
  const paused = hoverPause || tabHidden;

  const remainingRef = useRef(0);
  const startedRef = useRef(0);
  const timerRef = useRef(null);
  const swapRef = useRef(null);

  const advance = useCallback(() => {
    const goNext = () => {
      posRef.current += 1;
      if (posRef.current >= orderRef.current.length) {
        const last = orderRef.current[orderRef.current.length - 1];
        orderRef.current = shuffledOrder(list.length, last);
        posRef.current = 0;
      }
      setIdx(orderRef.current[posRef.current]);
      setCycle(c => c + 1);
      setShown(true);
    };
    if (reduceMotion) { goNext(); return; }
    setShown(false);
    swapRef.current = setTimeout(goNext, 450);
  }, [list.length, reduceMotion]);

  useEffect(() => {
    remainingRef.current = readMs(list[idx]);
  }, [idx, cycle, list]);

  useEffect(() => {
    if (list.length < 2 || paused || !shown) return undefined;
    startedRef.current = performance.now();
    timerRef.current = setTimeout(advance, remainingRef.current);
    return () => {
      clearTimeout(timerRef.current);
      remainingRef.current = Math.max(0, remainingRef.current - (performance.now() - startedRef.current));
    };
  }, [idx, cycle, paused, shown, advance, list.length]);

  useEffect(() => {
    const onVisibility = () => setTabHidden(document.visibilityState === "hidden");
    document.addEventListener("visibilitychange", onVisibility);
    return () => document.removeEventListener("visibilitychange", onVisibility);
  }, []);

  useEffect(() => () => {
    clearTimeout(timerRef.current);
    clearTimeout(swapRef.current);
  }, []);

  if (!list.length) return null;
  const text = list[idx] || list[0];

  return (
    <figure
      className="lg-quote-panel"
      aria-label={`Quotes from ${author.name}`}
      onMouseEnter={() => setHoverPause(true)}
      onMouseLeave={() => setHoverPause(false)}
    >
      <blockquote className={"lg-quote" + (shown ? "" : " is-out")} data-quote-index={idx}>
        <p>{"“"}{text}{"”"}</p>
      </blockquote>
      <figcaption className="lg-quote-by">
        <span className="lg-quote-name">{author.name}</span>
        <span className="lg-quote-title">{author.title}</span>
      </figcaption>
      {list.length > 1 && !reduceMotion && (
        <div className="lg-progress" aria-hidden="true">
          <div
            key={cycle}
            className="lg-progress-fill"
            style={{
              animationDuration: `${readMs(text)}ms`,
              animationPlayState: paused || !shown ? "paused" : "running",
            }}
          />
        </div>
      )}
    </figure>
  );
}

const MIN_PASSWORD = 12;
const MAX_PASSWORD = 72; // the most the server will accept
const RESEND_SECONDS = 60; // the server allows one email per address per minute

const OFFLINE = "Can't reach the server. Check your connection and try again.";

const NOTICE_TEXT = {
  "link-expired": "That email link has expired or was already used. Request a new one below.",
  "link-failed": "That email link didn't work. Request a new one below.",
};

function errInfo(err) {
  return {
    status: err && typeof err.status === "number" ? err.status : 0,
    code: (err && err.code) || "",
    message: (err && err.message) || "",
  };
}
const isOffline = e => e.status === 0 || e.status === 502 || e.status === 503 || e.status === 504;

function signInMessage(err) {
  const e = errInfo(err);
  if (e.code === "email_not_confirmed") return "Confirm your email first. Use the link we sent you, or choose Create a new account to get a new one.";
  if (e.status === 429 || e.code === "over_request_rate_limit") return "Too many attempts. Wait a few minutes, then try again.";
  if (e.code === "user_banned") return "This account is turned off. Contact the Gate Log admin.";
  if (isOffline(e)) return OFFLINE;
  if (e.status >= 500) return "Sign in isn't working right now. Try again in a few minutes.";
  return "Email or password is incorrect.";
}

function linkMessage(err) {
  const e = errInfo(err);
  if (e.code === "over_email_send_rate_limit" || e.status === 429) return "Too many emails were requested just now. Wait a few minutes, then try again.";
  if (e.code === "email_address_not_authorized") return "Email delivery isn't switched on for this address yet. Contact the Gate Log admin.";
  if (["signup_disabled", "otp_disabled", "email_provider_disabled"].includes(e.code)) return "New accounts are turned off right now. Contact the Gate Log admin.";
  if (e.code === "email_address_invalid" || e.code === "validation_failed") return "That email address isn't valid.";
  if (isOffline(e)) return OFFLINE;
  return "The email couldn't be sent right now. Try again in a few minutes.";
}

function passwordMessage(err) {
  const e = errInfo(err);
  if (e.code === "weak_password" || (err && Array.isArray(err.reasons))) {
    const reasons = (err && err.reasons) || [];
    if (reasons.includes("pwned")) return "That password has shown up in a known data breach. Choose a different one.";
    if (reasons.includes("characters")) return "Use a mix of upper and lower case letters, numbers, and symbols.";
    if (/at least \d+ characters/i.test(e.message)) return e.message;
    return "That password is too easy to guess. Choose a different one.";
  }
  if (e.code === "same_password") return "That's already your password. Choose a different one.";
  if (["reauthentication_needed", "session_not_found", "session_expired", "bad_jwt"].includes(e.code) || e.status === 401 || e.status === 403) {
    return "This link has timed out. Sign out and request a new one.";
  }
  if (e.status === 429) return "Too many attempts. Wait a few minutes, then try again.";
  if (isOffline(e)) return OFFLINE;
  return "The password couldn't be saved. Try again in a few minutes.";
}

function useCountdown() {
  const [left, setLeft] = useState(0);
  useEffect(() => {
    if (left <= 0) return undefined;
    const t = setTimeout(() => setLeft(s => s - 1), 1000);
    return () => clearTimeout(t);
  }, [left]);
  return [left, setLeft];
}

function Card({ onSubmit, title, children }) {
  const inner = (
    <>
      <div className="lg-card-accent" />
      <h1 id="lg-title" className="lg-card-title">{title}</h1>
      {children}
    </>
  );
  if (!onSubmit) return <section className="lg-card" aria-labelledby="lg-title">{inner}</section>;
  return (
    <form
      className="lg-card"
      aria-labelledby="lg-title"
      noValidate
      onSubmit={e => { e.preventDefault(); onSubmit(); }}
    >
      {inner}
    </form>
  );
}

function EmailInput({ value, onChange, invalid, autoComplete, autoFocus = true }) {
  return (
    <input
      id="lg-email"
      className="lg-input lg-plain"
      type="email"
      name="email"
      inputMode="email"
      autoComplete={autoComplete}
      autoCapitalize="off"
      autoCorrect="off"
      spellCheck={false}
      placeholder="name@mmrgrp.com"
      value={value}
      onChange={e => onChange(e.target.value)}
      autoFocus={autoFocus}
      aria-invalid={invalid ? "true" : "false"}
      aria-describedby={invalid ? "lg-error" : undefined}
    />
  );
}

function ErrorLine({ text }) {
  if (!text) return null;
  return <div id="lg-error" className="lg-error" role="alert">{text}</div>;
}

// Sign in, create a new account, forgot password, and the check your email view.
function AccountForms({ initialNotice }) {
  const [view, setView] = useState("signin");
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [showPw, setShowPw] = useState(false);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState({ text: "", field: "" });
  const [notice, setNotice] = useState(initialNotice || null);
  const [sent, setSent] = useState({ kind: "", to: "" });
  const [cooldown, setCooldown] = useCountdown();

  const fail = (text, field = "") => setError({ text, field });
  const clearError = () => { if (error.text) setError({ text: "", field: "" }); };
  const go = next => {
    setView(next);
    setError({ text: "", field: "" });
    setNotice(null);
    setShowPw(false);
  };
  const onEmail = v => { setEmail(v); clearError(); };

  const signIn = async () => {
    if (busy) return;
    const addr = normalizeEmail(email);
    if (!addr) { fail("Enter your work email.", "email"); return; }
    if (!isWorkEmail(addr)) { fail("Use your @mmrgrp.com work email.", "email"); return; }
    if (!password) { fail("Enter your password.", "password"); return; }
    setBusy(true);
    setError({ text: "", field: "" });
    try {
      const { error: err } = await supabase.auth.signInWithPassword({ email: addr, password });
      if (err) fail(signInMessage(err), "password");
      else setPassword("");
    } catch (_) {
      fail(OFFLINE, "password");
    } finally {
      setBusy(false);
    }
  };

  const sendLink = async kind => {
    if (busy) return;
    const fromSent = view === "sent";
    const addr = normalizeEmail(email);
    if (!addr) { fail("Enter your work email.", "email"); return; }
    if (!isWorkEmail(addr)) { fail("Use your @mmrgrp.com work email. Other addresses can't have an account.", "email"); return; }
    setBusy(true);
    setError({ text: "", field: "" });
    try {
      const { error: err } = kind === "register"
        ? await supabase.auth.signInWithOtp({
            email: addr,
            options: { shouldCreateUser: true, emailRedirectTo: siteAddress() },
          })
        : await supabase.auth.resetPasswordForEmail(addr, { redirectTo: siteAddress() });
      if (err) {
        fail(linkMessage(err), fromSent ? "" : "email");
      } else {
        setEmail(addr);
        setSent({ kind, to: addr });
        go("sent");
        setCooldown(RESEND_SECONDS);
      }
    } catch (_) {
      fail(OFFLINE, fromSent ? "" : "email");
    } finally {
      setBusy(false);
    }
  };

  if (view === "register" || view === "forgot") {
    const register = view === "register";
    return (
      <Card title={register ? "Create your account" : "Reset your password"} onSubmit={() => sendLink(view)}>
        <p className="lg-text">
          {register
            ? "Use your @mmrgrp.com work email. We'll send you a link to confirm it, then you'll choose your own password."
            : "Enter your work email and we'll send you a link to choose a new password."}
        </p>
        <label className="lg-label" htmlFor="lg-email">Work email</label>
        <EmailInput value={email} onChange={onEmail} invalid={error.field === "email"} autoComplete="email" />
        <ErrorLine text={error.text} />
        <button type="submit" className="lg-submit" disabled={busy}>
          {busy ? "Sending…" : register ? "Send confirmation link" : "Send reset link"}
        </button>
        <button type="button" className="lg-back" onClick={() => go("signin")}>Back to sign in</button>
      </Card>
    );
  }

  if (view === "sent") {
    const register = sent.kind === "register";
    return (
      <Card title="Check your email">
        <p className="lg-text" role="status">
          {register ? (
            <>We sent a link to <strong className="lg-addr">{sent.to}</strong>. Open it on this computer or your phone to confirm your email, then create your password.</>
          ) : (
            <>If <strong className="lg-addr">{sent.to}</strong> has a Gate Log account, a reset link is on its way. Open it on this computer or your phone to choose a new password.</>
          )}
        </p>
        <p className="lg-hint">Only the newest link works. If nothing arrives within a few minutes, check your junk or spam folder.</p>
        <ErrorLine text={error.text} />
        <button
          type="button"
          className="lg-secondary lg-gap"
          disabled={busy || cooldown > 0}
          onClick={() => sendLink(sent.kind)}
        >
          {busy ? "Sending…" : cooldown > 0 ? `Send another link in ${cooldown}s` : "Send another link"}
        </button>
        <button type="button" className="lg-back" onClick={() => go("signin")}>Back to sign in</button>
      </Card>
    );
  }

  return (
    <Card title="Sign in" onSubmit={signIn}>
      {notice && NOTICE_TEXT[notice] && <div className="lg-callout" role="status">{NOTICE_TEXT[notice]}</div>}
      <label className="lg-label" htmlFor="lg-email">Work email</label>
      <EmailInput value={email} onChange={onEmail} invalid={error.field === "email"} autoComplete="username" />
      <div className="lg-label-row">
        <label className="lg-label" htmlFor="lg-password">Password</label>
        <button type="button" className="lg-link" onClick={() => go("forgot")}>Forgot password?</button>
      </div>
      <div className="lg-pw-wrap">
        <input
          id="lg-password"
          className="lg-input"
          type={showPw ? "text" : "password"}
          name="password"
          autoComplete="current-password"
          autoCapitalize="off"
          autoCorrect="off"
          spellCheck={false}
          value={password}
          onChange={e => { setPassword(e.target.value); clearError(); }}
          aria-invalid={error.field === "password" ? "true" : "false"}
          aria-describedby={error.field === "password" ? "lg-error" : undefined}
        />
        <button
          type="button"
          className="lg-show"
          onClick={() => setShowPw(v => !v)}
          aria-pressed={showPw}
          aria-label={showPw ? "Hide password" : "Show password"}
        >{showPw ? "Hide" : "Show"}</button>
      </div>
      <ErrorLine text={error.text} />
      <button type="submit" className="lg-submit" disabled={busy}>
        {busy ? "Signing in…" : "Sign in"}
      </button>
      <div className="lg-divider"><span>New to the Gate Log?</span></div>
      <button type="button" className="lg-secondary" onClick={() => go("register")}>Create a new account</button>
    </Card>
  );
}

// Shown after someone arrives from an email link: a new account chooses its
// password here, and a reset sets a new one.
function PasswordStep({ mode, email, onSaved, onSignOut }) {
  const creating = mode !== "reset";
  const [pw1, setPw1] = useState("");
  const [pw2, setPw2] = useState("");
  const [show, setShow] = useState(false);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState({ text: "", field: "" });

  const fail = (text, field = "") => setError({ text, field });
  const clearError = () => { if (error.text) setError({ text: "", field: "" }); };

  const save = async () => {
    if (busy) return;
    const addr = normalizeEmail(email);
    const name = addr.split("@")[0];
    const typed = pw1.trim().toLowerCase();
    if (!pw1) { fail("Choose a password.", "pw1"); return; }
    if (pw1.length < MIN_PASSWORD) { fail(`Use at least ${MIN_PASSWORD} characters.`, "pw1"); return; }
    if (pw1.length > MAX_PASSWORD) { fail(`Use ${MAX_PASSWORD} characters or fewer.`, "pw1"); return; }
    if (typed === addr || (name && typed === name)) { fail("Don't use your email address as your password.", "pw1"); return; }
    if (pw1 !== pw2) { fail("The two passwords don't match.", "pw2"); return; }
    setBusy(true);
    setError({ text: "", field: "" });
    try {
      const { error: err } = await supabase.auth.updateUser({ password: pw1, data: { password_set: true } });
      if (err) {
        fail(passwordMessage(err), "pw1");
        setBusy(false);
        return;
      }
      setPw1("");
      setPw2("");
      onSaved();
    } catch (_) {
      fail(OFFLINE, "pw1");
      setBusy(false);
    }
  };

  const describe = field => [field === "pw1" ? "lg-pw-hint" : null, error.field === field ? "lg-error" : null]
    .filter(Boolean).join(" ") || undefined;

  return (
    <Card title={creating ? "Create your password" : "Set a new password"} onSubmit={save}>
      <p className="lg-text">
        {creating ? "Your email is confirmed. Choose a password for " : "Choose a new password for "}
        <strong className="lg-addr">{email}</strong>.
      </p>
      <input
        className="lg-offscreen"
        type="email"
        name="username"
        autoComplete="username"
        value={email}
        readOnly
        tabIndex={-1}
        aria-hidden="true"
      />
      <label className="lg-label" htmlFor="lg-new-password">New password</label>
      <div className="lg-pw-wrap">
        <input
          id="lg-new-password"
          className="lg-input"
          type={show ? "text" : "password"}
          name="new-password"
          autoComplete="new-password"
          autoCapitalize="off"
          autoCorrect="off"
          spellCheck={false}
          value={pw1}
          onChange={e => { setPw1(e.target.value); clearError(); }}
          autoFocus
          aria-invalid={error.field === "pw1" ? "true" : "false"}
          aria-describedby={describe("pw1")}
        />
        <button
          type="button"
          className="lg-show"
          onClick={() => setShow(v => !v)}
          aria-pressed={show}
          aria-label={show ? "Hide passwords" : "Show passwords"}
        >{show ? "Hide" : "Show"}</button>
      </div>
      <div id="lg-pw-hint" className="lg-hint">At least {MIN_PASSWORD} characters. A short phrase you'll remember works well.</div>
      <label className="lg-label lg-gap" htmlFor="lg-confirm-password">Confirm password</label>
      <input
        id="lg-confirm-password"
        className="lg-input lg-plain"
        type={show ? "text" : "password"}
        name="confirm-password"
        autoComplete="new-password"
        autoCapitalize="off"
        autoCorrect="off"
        spellCheck={false}
        value={pw2}
        onChange={e => { setPw2(e.target.value); clearError(); }}
        aria-invalid={error.field === "pw2" ? "true" : "false"}
        aria-describedby={describe("pw2")}
      />
      <ErrorLine text={error.text} />
      <button type="submit" className="lg-submit" disabled={busy}>
        {busy ? "Saving…" : creating ? "Create password" : "Save new password"}
      </button>
      <button type="button" className="lg-back" onClick={onSignOut}>Sign out</button>
    </Card>
  );
}

export default function LoginScreen({ step = null, email = "", notice = null, onPasswordSaved, onSignOut }) {
  return (
    <div className="lg-root">
      <style>{LOGIN_CSS}</style>
      <SiteBackdrop />
      <main className="lg-main">
        <header className="lg-brand">
          <MmrLogo height={52} />
          <div className="lg-system">CP2 GATE LOG</div>
          <div className="lg-tagline">Badge hours for the CP2 LNG field workforce</div>
        </header>

        {step
          ? <PasswordStep mode={step} email={email} onSaved={onPasswordSaved} onSignOut={onSignOut} />
          : <AccountForms initialNotice={notice} />}

        <QuoteRotator quotes={QUOTES} author={QUOTE_AUTHOR} />

        <p className="lg-notice">Authorized MMR personnel only.</p>
      </main>
    </div>
  );
}

const LOGIN_CSS = `
@import url('https://fonts.googleapis.com/css2?family=DM+Mono:wght@400;500&family=DM+Sans:wght@400;500;600;700&family=Source+Serif+4:ital,opsz,wght@1,8..60,400;1,8..60,500&display=swap');

.lg-root {
  position: relative;
  min-height: 100vh;
  min-height: 100dvh;
  overflow: hidden;
  color: #e9e6e1;
  font-family: 'DM Sans', 'Segoe UI', system-ui, sans-serif;
  background:
    radial-gradient(1100px 520px at 72% 100%, rgba(168,49,66,0.30), transparent 70%),
    radial-gradient(700px 380px at 81% 70%, rgba(248,199,0,0.07), transparent 70%),
    linear-gradient(180deg, #07080b 0%, #0e0a0d 50%, #1a0b10 100%);
}
.lg-root *, .lg-root *::before, .lg-root *::after { box-sizing: border-box; }
.lg-backdrop { position: absolute; inset: 0; pointer-events: none; }
.lg-grid {
  position: absolute; inset: 0; width: 100%; height: 100%;
  -webkit-mask-image: linear-gradient(to bottom, transparent 0%, #000 28%, #000 72%, transparent 100%);
  mask-image: linear-gradient(to bottom, transparent 0%, #000 28%, #000 72%, transparent 100%);
}
.lg-skyline {
  position: absolute; left: 0; right: 0; bottom: 0;
  width: 100%; height: 46vh; min-height: 260px; max-height: 540px;
}
.lg-main {
  position: relative; z-index: 1;
  min-height: 100vh; min-height: 100dvh;
  display: flex; flex-direction: column; align-items: center;
  padding: max(7vh, 32px) 16px 48px;
}
.lg-brand { display: flex; flex-direction: column; align-items: center; text-align: center; color: #fff; margin-bottom: 22px; }
.lg-system {
  margin-top: 14px; padding-left: 0.34em;
  font-family: 'DM Mono', ui-monospace, monospace; font-size: 13px; letter-spacing: 0.34em; color: #f8c700;
}
.lg-tagline { margin-top: 6px; font-size: 13px; color: #a9a39c; }
.lg-card {
  position: relative; width: 100%; max-width: 400px; overflow: hidden;
  background: rgba(13,10,12,0.80);
  border: 1px solid rgba(255,255,255,0.08); border-radius: 14px;
  padding: 28px 24px 24px;
  box-shadow: 0 24px 60px rgba(0,0,0,0.45);
  -webkit-backdrop-filter: blur(10px); backdrop-filter: blur(10px);
}
.lg-card-accent { position: absolute; top: 0; left: 0; right: 0; height: 3px; background: linear-gradient(90deg, #8a2432, #a83142 55%, #f8c700); }
.lg-card-title { margin: 0 0 18px; font-size: 18px; font-weight: 700; color: #f4f1ec; }
.lg-offscreen { position: absolute; left: -9999px; width: 1px; height: 1px; opacity: 0; pointer-events: none; }
.lg-label {
  display: block; margin-bottom: 8px;
  font-family: 'DM Mono', ui-monospace, monospace; font-size: 11px; letter-spacing: 0.14em; text-transform: uppercase; color: #a9a39c;
}
.lg-pw-wrap { position: relative; }
.lg-input {
  width: 100%; padding: 12px 70px 12px 14px;
  background: #0a0809; color: #f4f1ec;
  border: 1px solid #3a2a2e; border-radius: 8px; outline: none;
  font-family: 'DM Mono', ui-monospace, monospace; font-size: 16px;
  transition: border-color 0.15s, box-shadow 0.15s;
}
.lg-input:focus { border-color: #f8c700; box-shadow: 0 0 0 3px rgba(248,199,0,0.18); }
.lg-input[aria-invalid="true"] { border-color: #c2434f; }
.lg-show {
  position: absolute; right: 6px; top: 50%; transform: translateY(-50%);
  background: none; border: none; border-radius: 6px; padding: 6px 10px; cursor: pointer;
  font-family: 'DM Mono', ui-monospace, monospace; font-size: 12px; color: #a9a39c;
}
.lg-show:hover { color: #f4f1ec; }
.lg-show:focus-visible { color: #f4f1ec; outline: 2px solid rgba(248,199,0,0.45); }
.lg-error { margin-top: 10px; font-size: 13px; color: #ff8f8f; }
.lg-submit {
  margin-top: 16px; width: 100%; padding: 13px 16px; cursor: pointer;
  background: #8a2432; color: #fff; border: none; border-radius: 8px;
  font-family: inherit; font-size: 15px; font-weight: 600; letter-spacing: 0.02em;
  transition: background 0.15s, transform 0.05s;
}
.lg-submit:hover { background: #a83142; }
.lg-submit:active { transform: translateY(1px); }
.lg-submit:focus-visible { outline: 2px solid #f8c700; outline-offset: 2px; }
.lg-submit:disabled { opacity: 0.7; cursor: default; }
.lg-quote-panel {
  width: 100%; max-width: 400px; margin: 26px 0 0;
  padding: 2px 4px 0 18px; border-left: 2px solid #f8c700;
}
.lg-quote { margin: 0; transition: opacity 0.45s ease, transform 0.45s ease; }
.lg-quote.is-out { opacity: 0; transform: translateY(4px); }
.lg-quote p {
  margin: 0; font-family: 'Source Serif 4', Georgia, serif; font-style: italic;
  font-size: 18px; line-height: 1.5; color: #ece6dc;
}
.lg-quote-by { margin-top: 10px; display: flex; flex-wrap: wrap; align-items: baseline; gap: 4px 10px; }
.lg-quote-name {
  font-family: 'DM Mono', ui-monospace, monospace; font-size: 11px; letter-spacing: 0.2em;
  text-transform: uppercase; color: #f8c700;
}
.lg-quote-title { font-size: 12px; color: #a9a39c; }
.lg-progress { margin-top: 12px; height: 1px; overflow: hidden; background: rgba(255,255,255,0.08); }
.lg-progress-fill {
  height: 100%; background: rgba(248,199,0,0.7);
  transform-origin: left center; transform: scaleX(0);
  animation-name: lgFill; animation-timing-function: linear; animation-fill-mode: forwards;
}
@keyframes lgFill { from { transform: scaleX(0); } to { transform: scaleX(1); } }
.lg-notice { margin: 26px 0 0; font-size: 11px; letter-spacing: 0.04em; color: rgba(233,230,225,0.45); }
.lg-titleblock {
  position: absolute; right: 20px; bottom: 20px; z-index: 1; min-width: 250px;
  font-family: 'DM Mono', ui-monospace, monospace; font-size: 10px; letter-spacing: 0.12em; text-transform: uppercase;
  color: rgba(233,230,225,0.55);
  border: 1px solid rgba(255,255,255,0.14); background: rgba(8,9,12,0.6);
}
.lg-tb-row { display: flex; justify-content: space-between; gap: 16px; padding: 5px 10px; border-top: 1px solid rgba(255,255,255,0.08); }
.lg-tb-row:first-child { border-top: none; }
.lg-tb-head { justify-content: center; color: rgba(248,199,0,0.8); }
.lg-tb-row span:first-child { color: rgba(233,230,225,0.35); }
.lg-flame { transform-box: fill-box; transform-origin: 50% 100%; animation: lgFlicker 2.8s ease-in-out infinite; }
.lg-flare-glow { animation: lgGlow 2.8s ease-in-out infinite; }
.lg-beacon { animation: lgBlink 2.4s steps(1, end) infinite; }
.lg-beacon.b2 { animation-delay: 1.2s; }
@keyframes lgFlicker { 0%, 100% { transform: scale(1, 1); } 35% { transform: scale(0.94, 1.1); } 70% { transform: scale(1.04, 0.93); } }
@keyframes lgGlow { 0%, 100% { opacity: 0.5; } 35% { opacity: 0.62; } 70% { opacity: 0.44; } }
@keyframes lgBlink { 0%, 55% { opacity: 0.95; } 56%, 100% { opacity: 0.12; } }
.lg-label-row { display: flex; align-items: baseline; justify-content: space-between; gap: 12px; margin-top: 16px; }
.lg-gap { margin-top: 16px; }
.lg-plain { padding-right: 14px; }
.lg-text { margin: -6px 0 18px; font-size: 14px; line-height: 1.55; color: #c9c3bb; }
.lg-addr { color: #fff; font-weight: 600; word-break: break-all; }
.lg-hint { margin-top: 8px; font-size: 12px; line-height: 1.45; color: #8f8982; }
.lg-callout {
  margin: -4px 0 18px; padding: 10px 12px; border-radius: 8px;
  font-size: 13px; line-height: 1.45; color: #f3e3a1;
  background: rgba(248,199,0,0.08); border: 1px solid rgba(248,199,0,0.32);
}
.lg-link {
  background: none; border: none; padding: 0; margin-bottom: 8px; cursor: pointer;
  font-family: inherit; font-size: 13px; color: #e8bf2a;
}
.lg-link:hover { color: #f8c700; text-decoration: underline; }
.lg-link:focus-visible { outline: 2px solid rgba(248,199,0,0.45); outline-offset: 2px; border-radius: 4px; }
.lg-divider { display: flex; align-items: center; gap: 12px; margin: 22px 0 14px; font-size: 12px; color: #8f8982; }
.lg-divider::before, .lg-divider::after { content: ""; flex: 1; height: 1px; background: rgba(255,255,255,0.10); }
.lg-secondary {
  width: 100%; padding: 12px 16px; cursor: pointer;
  background: transparent; color: #f4f1ec;
  border: 1px solid rgba(248,199,0,0.45); border-radius: 8px;
  font-family: inherit; font-size: 15px; font-weight: 600;
  transition: background 0.15s, border-color 0.15s;
}
.lg-secondary:hover:not(:disabled) { background: rgba(248,199,0,0.08); border-color: #f8c700; }
.lg-secondary:focus-visible { outline: 2px solid #f8c700; outline-offset: 2px; }
.lg-secondary:disabled { opacity: 0.55; cursor: default; }
.lg-back {
  display: block; margin: 14px auto 0; padding: 6px 8px; cursor: pointer;
  background: none; border: none; border-radius: 4px;
  font-family: inherit; font-size: 13px; color: #a9a39c;
}
.lg-back:hover { color: #f4f1ec; text-decoration: underline; }
.lg-back:focus-visible { outline: 2px solid rgba(248,199,0,0.45); }
@media (max-width: 900px) { .lg-titleblock { display: none; } }
@media (max-width: 480px) {
  .lg-card { padding: 24px 18px 20px; }
  .lg-quote p { font-size: 16px; }
  .lg-skyline { height: 34vh; }
}
@media (prefers-reduced-motion: reduce) {
  .lg-flame, .lg-flare-glow, .lg-beacon { animation: none; }
  .lg-quote { transition: none; }
}
`;
