export function Footer() {
  return (
    <footer className="border-t border-border/50 bg-card/30 backdrop-blur-md py-6 mt-auto hidden md:block">
      <div className="container mx-auto px-4 flex flex-col md:flex-row items-center justify-between gap-4 text-xs text-muted-foreground">
        <div>
          © {new Date().getFullYear()} Tableau de bord CRC. SRM-SM. Tous droits réservés.
        </div>
        <div className="flex items-center gap-1.5 font-medium text-foreground">
          <span>Développé par</span>
                <span className="text-primary font-extrabold uppercase tracking-wide">
                Service Informatique DPIA
                </span>
        </div>
      </div>
    </footer>
  );
}
