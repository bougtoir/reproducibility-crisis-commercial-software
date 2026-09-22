# Agent reconstruction protocol

**DRAFT — NOT FROZEN. No authorized main-run configuration exists yet.**

## Proposed pilot envelope

Per slot: four vCPU, 16 GiB RAM, 100 GiB scratch, no GPU, 120 active wall minutes,
250,000 provider-accounted input+output tokens, and 1,000 tool calls. These are
planning ceilings, not measured usage or demonstrated available platform settings.
Pilot must verify enforcement, provider accounting, total-elapsed cap, outage
clock policy, maximum monetary cost and availability of the same model/version.
An unverifiable model version is recorded as unknown and limits the claim.
No main run starts with null budget/stop rules in the harness contract.

Use three fixed slots, same target, package hash, prompt, permissions, dependency
policy, Internet access boundary and resource tier. If a second tier is essential,
assign it prospectively by paper inputs and use it identically within that paper.
The preparation audit agents are not experiment slots.

## Solver instructions to freeze verbatim after pilot

1. Inspect permitted publication and target; record observable ambiguities with
   source locations, not private reasoning traces.
2. Produce an independent implementation faithful to the single reported method.
   Do not consult the original implementation, other runs or forbidden tools.
3. Install only broker-approved general dependencies. Record exact versions.
4. Execute and debug code within the fixed budget. Fix coding/environment errors;
   do not vary analytical specifications, models, seeds or defaults to improve
   target agreement. Record each attempt and implementation revision.
5. Generate the target once the declared procedure works. Use only the frozen
   comparison criterion. Record absolute/component errors and the target-linked
   conclusion separately.
6. Stop at completed assessment, verified unavoidable barrier, budget limit,
   suspected contamination or controller stop. Do not silently restart.
7. Submit code, permitted outputs, execution records, assumptions, ambiguities,
   failure evidence and structured outcomes. Missing usage is unknown, never zero.

## Required records

Run IDs and paper IDs; model/version/mode and exposed settings; protocol/prompt/
package/context/tool-registry/image hashes; exact UTC start/stop; stop reason;
L1–L5; failure and attribution evidence; attempts/revisions/dependencies;
elapsed time, tokens and costs; access events and contamination status.

Implementation is observed executable code, not a persuasive report. Execution
requires saved stdout/stderr/status/output artifacts. Success requires an
independently verifiable comparison; conclusion agreement does not substitute for
numerical agreement. An independent blinded adjudicator resolves ambiguous scores
before blind freeze, without inspecting original author code.
