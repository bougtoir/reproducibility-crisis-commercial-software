# Information firewall

**DRAFT — FULL FIREWALL NOT QUALIFIED. Separate VMs alone are insufficient.**

`paper2.isolation` and `paper2.model_api` provide synthetic execution-isolation
and API smoke qualifications only. Neither authorizes pilot or main runs.
The current execution probe uses one CPU, 2 GiB RAM and 512 MiB work storage;
it does not validate the larger proposed scientific-run envelope. A model alias
without a provider-reported version remains explicitly version-unverified.

Allowed: publication text, methodological supplements including published
scientific pseudocode, vetted public inputs, necessary general references and
free/open-source documentation. Forbidden: original repository, source, notebook,
workflow, container, alternative online reimplementations, and author material
that reveals implementation rather than scientific specification.

Use separate custodian, controller/broker, solver and adjudicator roles. A
custodian quarantines mixed archives and serves only reviewed exact members with
archive-to-member provenance. Authors' implementation is not opened to prepare
blind outcomes. Unknown/unchecked material is denied, even at an allowed domain.
Preserve content hash, bytes, licence, acquisition conditions, reviewer and
decision for every allowed or blocked artifact. Templates are in `schemas/`.

Each solver requires a new conversation, dedicated clean execution environment,
no writable shared volumes, no cached browser or shell credentials, and no
ancestor/sibling results. Shared read-only vetted software and paper inputs are
permitted. Restrict BOTH execution networking and model-side web/Git/MCP/session
tools. Audit injected rules, skills, knowledge, attachments, initial context and
effective tool registry. A container network restriction does not demonstrate
that model-side retrieval tools are restricted.

No parent debugging, package enrichment or policy widening during a run. Requests
to the broker must be logged; allow only materials under the frozen policy that
can be made equivalently available to all slots without outcome leakage.
Record attempts, denials, bytes served, timestamps and immutable event hashes.
Never record private chain-of-thought. Prior model-training exposure cannot be
certified absent and is a limitation even with an enforced runtime boundary.

Qualification must demonstrate: forbidden canary artifact cannot be read by
shell/browser/model-side tools; archive member and transitive dependency controls;
no cross-run read/write; effective-context inventory; complete access telemetry;
fixed resource stops; and frozen-outcome-before-reveal enforcement. These checks
must themselves have evidence and adversarial negative cases. The present
Devin preparation workflow does not implement these controls and must not be
relabelled a blinded reconstruction experiment.

Contamination: stop, seal evidence, adjudicate independently. Suspected or unknown
firewall status is unresolved for the clean primary analysis; confirmed
contamination is excluded, retaining its fixed slot as missing. No silent
replacement or repeated attempts until success. The retained contaminated trace
is not supplied to another solver.
