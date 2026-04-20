import { ScrollView, StyleSheet, Text, View } from "react-native";
import { featuredSermon, rhythmCards, takeaways } from "../data/mock";
import { palette } from "../theme/palette";
import { spacing } from "../theme/spacing";
import { typography } from "../theme/typography";

export function HomeScreen() {
  return (
    <ScrollView style={styles.screen} contentContainerStyle={styles.content}>
      <Text style={styles.eyebrow}>SERMONLIFE</Text>
      <Text style={styles.title}>Live the sermon after Sunday.</Text>
      <Text style={styles.subtitle}>
        Capture what you heard, pray through it, and carry it into the week.
      </Text>

      <View style={styles.heroCard}>
        <Text style={styles.sectionLabel}>Featured Sermon</Text>
        <Text style={styles.cardTitle}>{featuredSermon.title}</Text>
        <Text style={styles.cardMeta}>
          {featuredSermon.speaker} · {featuredSermon.date}
        </Text>
        <Text style={styles.scripture}>{featuredSermon.scripture}</Text>
        <Text style={styles.series}>{featuredSermon.series}</Text>
      </View>

      <Text style={styles.sectionHeading}>This Week's Rhythm</Text>
      <View style={styles.row}>
        {rhythmCards.map((item) => (
          <View style={styles.metricCard} key={item.label}>
            <Text style={styles.metricValue}>{item.value}</Text>
            <Text style={styles.metricLabel}>{item.label}</Text>
          </View>
        ))}
      </View>

      <Text style={styles.sectionHeading}>Application Prompts</Text>
      <View style={styles.listCard}>
        {takeaways.map((item) => (
          <Text key={item} style={styles.listItem}>
            {`\u2022 ${item}`}
          </Text>
        ))}
      </View>
    </ScrollView>
  );
}

const styles = StyleSheet.create({
  screen: {
    flex: 1,
    backgroundColor: palette.canvas
  },
  content: {
    paddingHorizontal: spacing.lg,
    paddingTop: spacing.xl,
    paddingBottom: spacing.xxl
  },
  eyebrow: {
    ...typography.overline,
    color: palette.accent
  },
  title: {
    ...typography.display,
    color: palette.ink,
    marginTop: spacing.sm
  },
  subtitle: {
    ...typography.body,
    color: palette.muted,
    marginTop: spacing.md
  },
  heroCard: {
    backgroundColor: palette.card,
    borderRadius: 28,
    padding: spacing.lg,
    marginTop: spacing.xl,
    shadowColor: "#000000",
    shadowOpacity: 0.08,
    shadowRadius: 18,
    shadowOffset: { width: 0, height: 10 },
    elevation: 3
  },
  sectionLabel: {
    ...typography.overline,
    color: palette.muted
  },
  cardTitle: {
    ...typography.title,
    color: palette.ink,
    marginTop: spacing.sm
  },
  cardMeta: {
    ...typography.body,
    color: palette.muted,
    marginTop: spacing.xs
  },
  scripture: {
    ...typography.headline,
    color: palette.accent,
    marginTop: spacing.lg
  },
  series: {
    ...typography.body,
    color: palette.ink,
    marginTop: spacing.xs
  },
  sectionHeading: {
    ...typography.headline,
    color: palette.ink,
    marginTop: spacing.xl,
    marginBottom: spacing.md
  },
  row: {
    flexDirection: "row",
    gap: spacing.md
  },
  metricCard: {
    flex: 1,
    backgroundColor: palette.cardSoft,
    borderRadius: 22,
    padding: spacing.md
  },
  metricValue: {
    ...typography.title,
    color: palette.ink
  },
  metricLabel: {
    ...typography.caption,
    color: palette.muted,
    marginTop: spacing.xs
  },
  listCard: {
    backgroundColor: palette.card,
    borderRadius: 24,
    padding: spacing.lg
  },
  listItem: {
    ...typography.body,
    color: palette.ink,
    marginBottom: spacing.sm
  }
});

