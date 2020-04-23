addFixture <- function(dt, fix2) {
	rbind(dt, t(as.data.table(strsplit(fix2, "-"))), use.names = FALSE)
}