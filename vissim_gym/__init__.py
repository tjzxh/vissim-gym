from gym.envs.registration import register

register(
    id='vissim-v11',
    entry_point='vissim_gym.envs:VissimEnv',
)
