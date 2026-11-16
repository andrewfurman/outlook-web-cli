//! Help quality: every command and flag is documented, every command has examples,
//! and the top-level help is one screen with no duplicated command list.

use assert_cmd::Command;

// Black-box: these read the real binary's help output. The clap tree itself is walked by
// a unit test in src/main.rs.

fn run(args: &[&str]) -> (bool, String) {
    let dir = tempfile::tempdir().unwrap();
    let out = Command::new(assert_cmd::cargo::cargo_bin!("outlook-web-cli"))
        .env_clear()
        .env("HOME", dir.path())
        .env("OUTLOOK_WEB_CONFIG_DIR", dir.path())
        .args(args)
        .output()
        .unwrap();
    (
        out.status.success(),
        String::from_utf8_lossy(&out.stdout).into_owned(),
    )
}

/// Every command path, found by reading each help's "Commands:" list; the top level
/// uses the hand-written list, so its paths come from the known first words.
const TOP: &[&str] = &[
    "auth",
    "accounts",
    "status",
    "messages",
    "search",
    "read",
    "agenda",
    "freebusy",
    "attachments",
    "attachment-inspect",
    "draft-new",
    "draft-reply",
    "draft-delete",
];

fn auth_subs() -> Vec<String> {
    let (_, text) = run(&["auth", "--help"]);
    let mut subs = Vec::new();
    let mut in_cmds = false;
    for line in text.lines() {
        if line.starts_with("Commands:") {
            in_cmds = true;
        } else if in_cmds {
            match line.strip_prefix("  ") {
                Some(rest) if !rest.starts_with(' ') => {
                    let name = rest.split_whitespace().next().unwrap().to_string();
                    if name != "help" {
                        subs.push(name);
                    }
                }
                _ => {
                    if !line.trim().is_empty() {
                        continue;
                    }
                    break;
                }
            }
        }
    }
    subs
}

fn all_paths() -> Vec<Vec<String>> {
    let mut paths: Vec<Vec<String>> = TOP.iter().map(|c| vec![c.to_string()]).collect();
    for s in auth_subs() {
        paths.push(vec!["auth".into(), s]);
    }
    assert!(paths.len() >= 17, "auth subcommands not found: {paths:?}");
    paths
}

#[test]
fn top_level_help_has_examples_and_is_the_same_everywhere() {
    let (ok, long) = run(&["--help"]);
    assert!(ok);
    for args in [vec![], vec!["-h"], vec!["help"]] {
        let (ok, text) = run(&args);
        assert!(ok, "{args:?}");
        assert_eq!(text, long, "{args:?} differs from --help");
    }
    for needle in [
        "EXAMPLES",
        "FIRST-TIME SETUP",
        "--yes",
        "--help",
        "JSON",
        "nothing is ever sent",
    ] {
        assert!(long.contains(needle), "top-level help lacks {needle}");
    }
    assert!(long.lines().count() <= 60, "top-level help is too long");
    // No second list of the commands: each command's description appears once.
    assert_eq!(long.matches("Full-mailbox search").count(), 1);
    assert!(!long.contains("\nCommands:\n"));
    // Every command is listed, and every example is a real command.
    for c in TOP {
        assert!(
            long.contains(&format!("  {c} ")),
            "{c} missing from top help"
        );
    }
    for line in long.lines().filter(|l| l.starts_with("  outlook-web-cli ")) {
        let word = line.split_whitespace().nth(1).unwrap();
        assert!(TOP.contains(&word), "example uses unknown command: {line}");
    }
}

#[test]
fn every_command_has_help_with_examples_and_documented_flags() {
    for path in all_paths() {
        let name = path.join(" ");
        let mut long: Vec<&str> = path.iter().map(String::as_str).collect();
        long.push("--help");
        let (ok, text) = run(&long);
        assert!(ok, "{name} --help failed");
        assert!(text.contains("Examples:"), "{name} has no Examples section");
        let examples = text
            .lines()
            .filter(|l| l.trim_start().starts_with("outlook-web-cli ") && l.starts_with("  "))
            .count();
        assert!(examples >= 1, "{name} has no example lines");
        // -h, --help and `help <cmd>` print the same thing.
        let mut short: Vec<&str> = path.iter().map(String::as_str).collect();
        short.push("-h");
        assert_eq!(run(&short).1, text, "{name}: -h differs from --help");
        let mut via: Vec<&str> = vec!["help"];
        via.extend(path.iter().map(String::as_str));
        assert_eq!(run(&via).1, text, "help {name} differs from --help");
        // Every flag line carries a description.
        let mut flags = 0;
        let mut lines = text.lines().peekable();
        while let Some(line) = lines.next() {
            let t = line.trim_start();
            if !(line.starts_with("  ") && t.starts_with('-')) {
                continue;
            }
            flags += 1;
            // "  -h, --help   Print help": description on the same line after 2+ spaces,
            // or on the next, more indented line.
            let same_line = t.split("  ").skip(1).any(|p| !p.trim().is_empty());
            let next_indented = lines
                .peek()
                .is_some_and(|n| n.starts_with("          ") && !n.trim().is_empty());
            assert!(
                same_line || next_indented,
                "{name}: flag without help: {line}"
            );
        }
        assert!(flags >= 1, "{name} lists no flags");
    }
}
